"""Agent-friendly kwarg normalization for the one-call shape helpers.

Code generators (LLM-trained models among them) reliably spell keyword
arguments the way *other* libraries spell them — matplotlib's
``fontfamily``/``fontsize``/``ha``/``va``, CSS-flavored ``text-align``,
``colour`` — and die on ``TypeError: unexpected keyword argument``.  This
module absorbs that whole error class for the ergonomic helpers
(``add_text`` / ``add_equation`` / ``add_arrow``) in three layers:

1.  **Synonyms** — a fixed map from common spellings onto each canonical
    argument (``halign`` → ``align``, ``font_size`` → ``size_pt`` …).
2.  **Fuzzy matching** — a near-miss kwarg (``algn``, ``colr``) is resolved
    to its closest canonical or synonym when unambiguous.
3.  **Didactic errors** — anything still unknown raises a ``TypeError``
    that names the closest candidate and lists every accepted kwarg, so a
    model reading the traceback self-corrects in one step.

An alias may *substitute* for its canonical but never contradict it:
passing two different values for the same logical argument is a genuine
caller bug and still raises.
"""

from __future__ import annotations

import difflib

# Canonical kwarg -> spellings other ecosystems use for the same thing.
SYNONYMS: dict[str, tuple[str, ...]] = {
    "text": ("txt", "string", "content", "label", "caption", "value"),
    "font": ("font_family", "fontfamily", "font_name", "fontname", "typeface", "family", "face"),
    "size_pt": ("size", "font_size", "fontsize", "pt_size", "point_size"),
    "align": ("halign", "ha", "horizontal_align", "horizontal_alignment", "text_align", "text_alignment"),
    "anchor": ("valign", "va", "vertical_align", "vertical_alignment", "v_align"),
    "color": (
        "colour",
        "font_color",
        "font_colour",
        "text_color",
        "text_colour",
        "fg_color",
        "line_color",
        "stroke_color",
    ),
    "weight_pt": ("weight", "line_weight", "width_pt", "stroke_width"),
    "bold": ("font_bold", "is_bold"),
    "italic": ("font_italic", "is_italic", "oblique"),
    "word_wrap": ("wrap", "wrap_text", "text_wrap"),
    "margin_pt": ("margin", "padding", "padding_pt", "inset", "inset_pt"),
    "latex": ("tex", "formula", "equation", "expression", "math"),
    "left": ("x",),
    "top": ("y",),
    "width": ("w",),
    "height": ("h",),
    "start": ("begin", "start_point", "start_shape", "from", "source"),
    "end": ("to", "end_point", "end_shape", "target"),
}

_ALIAS_TO_CANONICAL: dict[str, str] = {
    alias: canonical
    for canonical, aliases in SYNONYMS.items()
    for alias in aliases
}

_FUZZY_CUTOFF = 0.85


def _did_you_mean(name: str, candidates: list[str], known: "list[str] | tuple[str, ...]") -> str | None:
    """Return the canonical arg *name* points at, or None.

    Near-misses that hit both a canonical and one of its own aliases
    (``algn`` → ``align`` and ``halign``) are NOT ambiguous — they agree —
    so candidates are collapsed to their canonical form before the
    ambiguity check.
    """
    if not candidates:
        return None
    matches = difflib.get_close_matches(name, candidates, n=3, cutoff=_FUZZY_CUTOFF)
    distinct = set()
    for m in matches:
        canonical = m if m in known else _ALIAS_TO_CANONICAL.get(m)
        if canonical in known:
            distinct.add(canonical)
    if len(distinct) == 1:
        return distinct.pop()
    return None


def absorb_agent_kwargs(
    method: str,
    kwargs: dict,
    canonical_names: "list[str] | tuple[str, ...]",
) -> dict:
    """Map alias / near-miss *kwargs* onto canonical names; raise else.

    *canonical_names* lists the kwargs the method actually understands
    (including any already-bound explicit parameters).  Returns a dict
    containing only canonical names.  A synonym that matches a canonical
    the caller already supplied with a *different* value raises
    ``TypeError``; equal values are fine.
    """
    known = list(canonical_names)
    alias_space = [a for a in _ALIAS_TO_CANONICAL if a not in known]
    candidates = known + alias_space

    resolved: dict = {}
    for name, value in kwargs.items():
        if name in known or name in resolved:
            if name in resolved and resolved[name] != value:
                raise TypeError(
                    f"{method}(): got {name!r} twice with different values "
                    f"({resolved[name]!r} vs {value!r})"
                )
            resolved[name] = value
            continue
        canonical = _ALIAS_TO_CANONICAL.get(name)
        if canonical is not None and canonical not in known:
            # The alias resolves to an argument this method doesn't take
            # (e.g. ``to=`` on add_text); fall through to the fuzzy/error
            # path rather than silently dropping the value.
            canonical = None
        if canonical is None:
            canonical = _did_you_mean(name, candidates, known)
            if canonical is None:
                raise TypeError(
                    f"{method}(): got an unexpected keyword argument {name!r}. "
                    f"Accepted: {', '.join(sorted(known))} "
                    f"(synonyms like font_family/halign/valign/colour are fine)"
                )
        if canonical in resolved and resolved[canonical] != value:
            raise TypeError(
                f"{method}(): {name!r} and the value already given for "
                f"{canonical!r} disagree ({value!r} vs {resolved[canonical]!r}); "
                f"pass one spelling only"
            )
        resolved[canonical] = value
    return resolved
