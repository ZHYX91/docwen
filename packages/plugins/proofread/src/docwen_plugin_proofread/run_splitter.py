"""Run splitting for precise comment anchoring in DOCX proofreading.

Word comments attach to an entire run.  When a proofreading error starts
or ends in the middle of a run (or spans multiple runs), we split at both
half-open range boundaries and anchor the comment to the complete run span.
"""

from __future__ import annotations

from copy import deepcopy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run


def plan_run_splits(paragraph, errors: list) -> list[int]:
    """Determine which character positions in *paragraph* need run splits.

    Analyses each error's ``start_pos`` and ``end_pos`` and returns a sorted list of
    unique positions where a run boundary is required.

    Returns an empty list when no splitting is needed (all error boundaries
    already fall on run boundaries).
    """
    if not paragraph.runs or not errors:
        return []

    split_positions: set[int] = set()
    for err in errors:
        for pos in (err.start_pos, err.end_pos):
            current = 0
            for run in paragraph.runs:
                run_len = len(run.text)
                if current < pos < current + run_len:
                    split_positions.add(pos)
                    break
                current += run_len
    return sorted(split_positions)


def ensure_run_at_position(paragraph, position: int) -> object:
    """Ensure there is a run boundary at the character *position*.

    If *position* falls inside an existing run, splits that run into two
    at the exact character offset.  The original run retains the text
    before the split point; a new run (cloned with formatting) receives
    the text after the split point.

    Returns the ``Run`` object that starts at *position*.
    """
    if not paragraph.runs:
        return paragraph.add_run("")

    current = 0
    for run in paragraph.runs:
        run_len = len(run.text)
        run_end = current + run_len

        if current == position:
            # Already at a run boundary — nothing to do
            return run

        if current < position < run_end:
            offset = position - current
            r_elem = run._r

            # ``Run.text`` flattens OOXML controls such as ``w:tab`` and
            # ``w:br`` to ``\t``/``\n``.  Re-emitting those characters in a
            # new ``w:t`` while retaining the original controls duplicates or
            # reorders user content.  Until splitting is token-aware, fail
            # closed for every run that is not composed solely of properties
            # and text nodes.  The caller will then decline to create an
            # inexact comment anchor, while the document remains byte-for-byte
            # unchanged at this run.
            allowed_children = {qn("w:rPr"), qn("w:t")}
            if any(child.tag not in allowed_children for child in r_elem.iterchildren()):
                return run

            # Collect the <w:t> children of this run
            t_elems = list(r_elem.iterchildren(qn("w:t")))
            if not t_elems:
                return run

            before_text = run.text[:offset]
            after_text = run.text[offset:]

            # ── Update the original run: keep only text before split ──
            for i, t in enumerate(t_elems):
                if i == 0:
                    t.text = before_text
                else:
                    r_elem.remove(t)

            # ── Create a new <w:r> for the text after the split ──
            new_r = OxmlElement("w:r")

            # Copy run properties (rPr, e.g. bold/italic/font) if present
            rPr = r_elem.find(qn("w:rPr"))
            if rPr is not None:
                new_r.append(deepcopy(rPr))

            # Create a new <w:t> for the second half
            new_t = OxmlElement("w:t")
            new_t.set(qn("xml:space"), "preserve")
            new_t.text = after_text
            new_r.append(new_t)

            # Insert the new run immediately after the original
            p_elem = paragraph._element
            idx = list(p_elem).index(r_elem)
            p_elem.insert(idx + 1, new_r)

            # Return a Run-like object wrapping the new element
            return Run(new_r, run._parent)  # type: ignore[arg-type]

        current = run_end

    # Position is beyond all existing runs — return the last run
    return paragraph.runs[-1] if paragraph.runs else paragraph.add_run("")


def runs_for_range(paragraph, start: int, end: int) -> list[Run]:
    """Return the run sequence that exactly covers half-open ``[start, end)``.

    Callers must first apply every position returned by :func:`plan_run_splits`.
    The function fails closed with an empty list if the requested range is
    invalid or either boundary still falls inside a run.
    """
    if start < 0 or end <= start or end > len(paragraph.text):
        return []

    selected: list[Run] = []
    selected_start: int | None = None
    selected_end: int | None = None
    current = 0
    for run in paragraph.runs:
        run_end = current + len(run.text)
        if run_end > start and current < end:
            if selected_start is None:
                selected_start = current
            selected.append(run)
            selected_end = run_end
        current = run_end

    if selected_start != start or selected_end != end:
        return []
    return selected
