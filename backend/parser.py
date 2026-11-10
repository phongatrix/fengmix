"""
parser.py - Parse .docx exam files into structured question data.

Supports Vietnamese exam format with sections (PHẦN I, II, III or Nhóm X),
multiple question types (trắc nghiệm, đúng sai, số lượng), and answer hints
embedded as "(Đáp án: X)" or "[Đáp án: X]" in the document.
"""

import re
from docx import Document
from io import BytesIO
from typing import Optional


# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# "Câu 1:" or "Câu 1." (with optional trailing space)
RE_QUESTION = re.compile(
    r'^[Cc]âu\s*(\d+)\s*[:.]\s*(.*)',
    re.UNICODE | re.IGNORECASE
)

# Option lines: "A. text", "B. text", etc.
RE_OPTION = re.compile(
    r'^([A-D])\.\s*(.*)',
    re.UNICODE
)

# Correct answer hint: (Đáp án: B) or [Đáp án: B]
RE_ANSWER_HINT = re.compile(
    r'[\[\(][Đđ]áp\s+án\s*:\s*([A-D])[\]\)]',
    re.UNICODE | re.IGNORECASE
)

# Section / group headers
RE_SECTION = re.compile(
    r'^(PH[ẦA]N\s+[IVX\d]+|NH[ÓO]M\s+\d+|PHẦN\s+[IVX\d]+)',
    re.UNICODE | re.IGNORECASE
)

# Metadata patterns
RE_DURATION = re.compile(r'(\d+)\s*phút', re.UNICODE | re.IGNORECASE)
RE_SUBJECT = re.compile(
    r'môn\s*(?:thi|kiểm tra)?:?\s*(.+)',
    re.UNICODE | re.IGNORECASE
)


def _clean(text: str) -> str:
    """Strip surrounding whitespace from a string."""
    return text.strip()


def _detect_question_type(group_name: str, question_text: str) -> str:
    """Heuristically detect question type from group name or question text."""
    combined = (group_name + " " + question_text).lower()
    if any(kw in combined for kw in ["đúng", "sai", "true", "false"]):
        return "dungsai"
    if any(kw in combined for kw in ["số lượng", "điền số", "tính"]):
        return "soluong"
    return "tracnghiem"


def _extract_metadata(paragraphs: list[str]) -> dict:
    """
    Try to extract metadata (title, subject, duration, school) from the
    first several non-empty lines of the document.
    """
    metadata = {
        "title": "",
        "subject": "",
        "duration": "",
        "school": "",
    }

    candidates = [p for p in paragraphs[:20] if p.strip()]

    for i, line in enumerate(candidates):
        # School name is usually the first bold/caps line
        if i == 0 and not metadata["school"]:
            metadata["school"] = _clean(line)

        # Duration
        m = RE_DURATION.search(line)
        if m and not metadata["duration"]:
            metadata["duration"] = m.group(1)

        # Subject
        m = RE_SUBJECT.search(line)
        if m and not metadata["subject"]:
            metadata["subject"] = _clean(m.group(1))

        # Title – lines containing "ĐỀ KIỂM TRA" or "ĐỀ THI"
        if re.search(r'[Đđ]ề\s*(kiểm tra|thi)', line, re.IGNORECASE) and not metadata["title"]:
            metadata["title"] = _clean(line)

    # Fallback title
    if not metadata["title"] and candidates:
        metadata["title"] = candidates[0]

    return metadata


def parse_docx(file_bytes: bytes) -> dict:
    """
    Parse a .docx exam file and return structured data.

    Returns:
        {
            "metadata": {
                "title": str,
                "subject": str,
                "duration": str,  # minutes
                "school": str,
            },
            "groups": [
                {
                    "name": str,
                    "questions": [
                        {
                            "num": int,
                            "text": str,
                            "options": {"A": str, "B": str, "C": str, "D": str},
                            "correct": "A"|"B"|"C"|"D",
                            "type": str,
                            "group_name": str,
                        },
                        ...
                    ]
                },
                ...
            ]
        }
    """
    doc = Document(BytesIO(file_bytes))

    # Collect all paragraph texts (preserve order)
    raw_paragraphs: list[str] = [p.text for p in doc.paragraphs]

    # Collect inline answer hints per paragraph index
    answer_hints: dict[int, str] = {}
    for idx, p in enumerate(doc.paragraphs):
        full_text = p.text
        m = RE_ANSWER_HINT.search(full_text)
        if m:
            answer_hints[idx] = m.group(1).upper()

    metadata = _extract_metadata(raw_paragraphs)

    # ------------------------------------------------------------------
    # State machine to build groups and questions
    # ------------------------------------------------------------------
    groups: list[dict] = []
    current_group_name = "PHẦN I"
    current_questions: list[dict] = []

    current_q: Optional[dict] = None
    last_option_key: Optional[str] = None

    def _finish_question():
        nonlocal current_q
        if current_q is None:
            return
        # Fill missing options with empty string
        for letter in "ABCD":
            current_q["options"].setdefault(letter, "")
        # Ensure correct answer is valid
        if current_q["correct"] not in current_q["options"]:
            current_q["correct"] = "A"
        current_q["type"] = _detect_question_type(current_group_name, current_q["text"])
        current_q["group_name"] = current_group_name
        current_questions.append(current_q)
        current_q = None

    def _finish_group():
        nonlocal current_questions, current_group_name
        if current_questions:
            groups.append({"name": current_group_name, "questions": list(current_questions)})
        current_questions = []

    for idx, raw in enumerate(raw_paragraphs):
        line = _clean(raw)
        if not line:
            continue

        # --- Section header detection ---
        if RE_SECTION.match(line):
            _finish_question()
            _finish_group()
            current_group_name = line
            continue

        # --- Question detection ---
        m = RE_QUESTION.match(line)
        if m:
            _finish_question()
            q_num = int(m.group(1))
            q_text_part = _clean(m.group(2))

            # Strip embedded answer hint from question text
            q_text_part = RE_ANSWER_HINT.sub("", q_text_part).strip()

            current_q = {
                "num": q_num,
                "text": q_text_part,
                "options": {},
                "correct": answer_hints.get(idx, "A"),
            }
            last_option_key = None
            continue

        # --- Option detection ---
        m = RE_OPTION.match(line)
        if m and current_q is not None:
            opt_key = m.group(1).upper()
            opt_text = _clean(m.group(2))
            # Strip answer hint from option text
            hint_match = RE_ANSWER_HINT.search(opt_text)
            if hint_match:
                current_q["correct"] = hint_match.group(1).upper()
                opt_text = RE_ANSWER_HINT.sub("", opt_text).strip()
            current_q["options"][opt_key] = opt_text
            last_option_key = opt_key
            continue

        # --- Continuation of previous option or question text ---
        if current_q is not None:
            # Strip answer hint if present
            hint_match = RE_ANSWER_HINT.search(line)
            if hint_match:
                current_q["correct"] = hint_match.group(1).upper()
                line = RE_ANSWER_HINT.sub("", line).strip()

            if last_option_key and last_option_key in current_q["options"]:
                # Continuation of last option
                current_q["options"][last_option_key] += " " + line
            else:
                # Continuation of question text
                current_q["text"] += " " + line

    # Flush last question and group
    _finish_question()
    _finish_group()

    # If no groups found, create a default one
    if not groups:
        groups = [{"name": "PHẦN I", "questions": []}]

    return {"metadata": metadata, "groups": groups}
