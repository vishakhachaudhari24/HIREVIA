"""
=============================================================
AI-ASSISTED RESUME SCREENING SYSTEM - scorer.py
=============================================================
Pure Python scoring module.
No AI calls here — only math.

Formula:
  Final Score = (0.40 × skill_match_score)
              + (0.20 × experience_relevance_score)
              + (0.20 × project_relevance_score)
              + (0.10 × domain_fit_score)
              + (0.10 × career_progression_score)

Returns a rounded integer (0–100).
=============================================================
"""


def calculate_final_score(evaluation: dict) -> int:
    """
    Calculate the weighted final match score from the AI evaluation dict.

    Parameters
    ----------
    evaluation : dict
        Must contain the five score keys returned by the AI:
          - skill_match_score          (weight 40%)
          - experience_relevance_score (weight 20%)
          - project_relevance_score    (weight 20%)
          - domain_fit_score           (weight 10%)
          - career_progression_score   (weight 10%)

    Returns
    -------
    int
        Weighted final score, clamped between 0 and 100.
    """

    # ── Safely read each score, defaulting to 0 if missing/invalid ──
    def safe_score(key: str) -> float:
        val = evaluation.get(key, 0)
        try:
            val = float(val)
        except (TypeError, ValueError):
            val = 0.0
        # Clamp individual scores to valid range
        return max(0.0, min(100.0, val))

    skill_match          = safe_score("skill_match_score")
    experience_relevance = safe_score("experience_relevance_score")
    project_relevance    = safe_score("project_relevance_score")
    domain_fit           = safe_score("domain_fit_score")
    career_progression   = safe_score("career_progression_score")

    # ── Apply weighted formula ───────────────────────────────────────
    final = (
        0.40 * skill_match
        + 0.20 * experience_relevance
        + 0.20 * project_relevance
        + 0.10 * domain_fit
        + 0.10 * career_progression
    )

    # ── Clamp total and return as integer ────────────────────────────
    return round(max(0.0, min(100.0, final)))