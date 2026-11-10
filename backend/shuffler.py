"""
shuffler.py - Deterministic exam shuffling by exam code (mã đề).

Supported shuffle modes:
  "ca_cau_va_dap_an" - shuffle both question order and answer choices
  "chi_cau_hoi"      - shuffle only question order, keep A/B/C/D in place
  "chi_dap_an"       - keep question order, shuffle only answer choices
  "theo_nhom"        - shuffle within each group, also shuffle group order
"""

import random
import copy
from typing import Any


LETTERS = ["A", "B", "C", "D"]


def _seed_for(ma_de: int) -> int:
    return ma_de * 31337 + 42


def _shuffle_answers(question: dict, rng: random.Random) -> dict:
    """
    Return a NEW question dict with shuffled answer options.
    Tracks which new letter now corresponds to the original correct answer.
    """
    q = copy.deepcopy(question)
    old_letters = list(LETTERS)
    new_letters = list(LETTERS)
    rng.shuffle(new_letters)

    # Build mapping: old_letter -> new_letter
    old_to_new: dict[str, str] = {old: new for old, new in zip(old_letters, new_letters)}

    # Rearrange options
    new_options: dict[str, str] = {}
    for old_l, new_l in old_to_new.items():
        new_options[new_l] = q["options"].get(old_l, "")

    q["options"] = new_options
    q["correct"] = old_to_new.get(q.get("correct", "A"), "A")
    return q


def _shuffle_questions(questions: list[dict], rng: random.Random) -> list[dict]:
    """Return a new list of questions in shuffled order, renumbered from 1."""
    shuffled = list(questions)
    rng.shuffle(shuffled)
    result = []
    for i, q in enumerate(shuffled, start=1):
        new_q = copy.deepcopy(q)
        new_q["num"] = i
        result.append(new_q)
    return result


def _build_answer_key(groups: list[dict]) -> dict[int, str]:
    """
    Build answer_key: {question_num (int): correct_letter (str)}
    Questions are numbered globally across all groups.
    """
    key: dict[int, str] = {}
    for group in groups:
        for q in group["questions"]:
            key[q["num"]] = q.get("correct", "A")
    return key


def _renumber_global(groups: list[dict]) -> list[dict]:
    """Renumber all questions globally (across groups) starting from 1."""
    counter = 1
    new_groups = []
    for group in groups:
        new_qs = []
        for q in group["questions"]:
            nq = copy.deepcopy(q)
            nq["num"] = counter
            counter += 1
            new_qs.append(nq)
        new_groups.append({**group, "questions": new_qs})
    return new_groups


def shuffle_exam(parsed_data: dict, ma_de: int, shuffle_mode: str) -> dict:
    """
    Shuffle exam data deterministically by exam code (mã đề).

    Args:
        parsed_data: Output of parser.parse_docx
        ma_de: Integer exam code (mã đề)
        shuffle_mode: One of "ca_cau_va_dap_an", "chi_cau_hoi",
                      "chi_dap_an", "theo_nhom"

    Returns:
        dict with keys: "metadata", "groups", "answer_key"
        answer_key: {question_num (int): correct_letter (str)}
    """
    rng = random.Random(_seed_for(ma_de))

    # Deep copy to avoid mutating the original
    data = copy.deepcopy(parsed_data)
    metadata = data.get("metadata", {})
    groups: list[dict] = data.get("groups", [])

    if shuffle_mode == "ca_cau_va_dap_an":
        # 1. Flatten all questions globally
        all_questions: list[dict] = []
        for group in groups:
            all_questions.extend(group["questions"])

        # 2. Shuffle both question order and answer choices
        rng.shuffle(all_questions)
        shuffled_qs: list[dict] = []
        for i, q in enumerate(all_questions, start=1):
            new_q = _shuffle_answers(q, rng)
            new_q["num"] = i
            shuffled_qs.append(new_q)

        # 3. Put all into one flat group preserving original group names
        new_groups = [{"name": groups[0]["name"] if groups else "PHẦN I", "questions": shuffled_qs}]

    elif shuffle_mode == "chi_cau_hoi":
        # Shuffle only question order globally (no answer permutation)
        all_questions = []
        for group in groups:
            all_questions.extend(copy.deepcopy(group["questions"]))

        rng.shuffle(all_questions)
        for i, q in enumerate(all_questions, start=1):
            q["num"] = i

        new_groups = [{"name": groups[0]["name"] if groups else "PHẦN I", "questions": all_questions}]

    elif shuffle_mode == "chi_dap_an":
        # Keep question order, shuffle only answer choices
        new_groups = []
        q_counter = 1
        for group in groups:
            new_qs = []
            for q in group["questions"]:
                new_q = _shuffle_answers(q, rng)
                new_q["num"] = q_counter
                q_counter += 1
                new_qs.append(new_q)
            new_groups.append({**group, "questions": new_qs})

    elif shuffle_mode == "theo_nhom":
        # Shuffle group order, then shuffle questions within each group,
        # and also shuffle answer choices within each question
        group_list = list(groups)
        rng.shuffle(group_list)

        new_groups = []
        q_counter = 1
        for group in group_list:
            qs = list(group["questions"])
            rng.shuffle(qs)
            new_qs = []
            for q in qs:
                new_q = _shuffle_answers(q, rng)
                new_q["num"] = q_counter
                q_counter += 1
                new_qs.append(new_q)
            new_groups.append({**group, "questions": new_qs})

    else:
        raise ValueError(f"Chế độ trộn không hợp lệ: '{shuffle_mode}'")

    answer_key = _build_answer_key(new_groups)

    return {
        "metadata": metadata,
        "groups": new_groups,
        "answer_key": answer_key,
    }
