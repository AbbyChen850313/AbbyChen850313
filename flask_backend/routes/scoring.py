"""
/api/scoring  — draft save, score submit, status queries.
"""

from __future__ import annotations

import logging

from flask import Blueprint, g, jsonify, request

from services.auth_service import require_auth, require_manager
from services.scoring_service import (
    calc_all,
    current_quarter,
    is_in_scoring_period,
)
from services.sheets_service import SheetsService

logger = logging.getLogger(__name__)
scoring_bp = Blueprint("scoring", __name__)


def _sheets(is_test: bool) -> SheetsService:
    return SheetsService(is_test=is_test)


# ── POST /api/scoring/draft ────────────────────────────────────────────────

@scoring_bp.route("/draft", methods=["POST"])
@require_manager
def save_draft():
    return _upsert_score(status="草稿")


# ── POST /api/scoring/submit ───────────────────────────────────────────────

@scoring_bp.route("/submit", methods=["POST"])
@require_manager
def submit_score():
    return _upsert_score(status="已送出")


def _upsert_score(status: str):
    """Shared logic for save_draft and submit_score."""
    session = g.session
    manager_name: str = session["name"]
    line_uid: str = session["lineUid"]
    is_test: bool = session.get("isTest", False)

    body = request.get_json(silent=True) or {}
    emp_name = (body.get("empName") or "").strip()
    section = (body.get("section") or "").strip()
    scores_raw: dict = body.get("scores") or {}
    special = float(body.get("special") or 0)
    note = (body.get("note") or "").strip()
    quarter = (body.get("quarter") or "").strip()

    if not emp_name or not section:
        return jsonify({"error": "缺少員工姓名或科別"}), 400

    sheets = _sheets(is_test)
    settings = sheets.get_settings()

    if not quarter:
        quarter = settings.get("當前季度") or current_quarter()

    if status == "已送出":
        # Validate scoring period
        if not is_in_scoring_period(settings):
            return jsonify({"error": "不在評分期間內，無法送出"}), 403

        # Validate all 6 items are filled
        missing = [f"item{i}" for i in range(1, 7) if not scores_raw.get(f"item{i}")]
        if missing:
            return jsonify({"error": f"評分項目未填完：{', '.join(missing)}"}), 400

    # Get manager weight for this section
    responsibilities = sheets.get_manager_responsibilities()
    weight = next(
        (r["weight"] for r in responsibilities if r["lineUid"] == line_uid and r["section"] == section),
        0.0,
    )

    calc = calc_all(scores_raw, special, weight)

    score_data = {
        "quarter": quarter,
        "managerName": manager_name,
        "empName": emp_name,
        "section": section,
        "weight": weight,
        "scores": scores_raw,
        "special": special,
        "note": note,
        "status": status,
        **calc,
    }

    sheets.upsert_score(score_data)

    return jsonify({
        "success": True,
        "status": status,
        **calc,
    })


# ── GET /api/scoring/my-scores ─────────────────────────────────────────────

@scoring_bp.route("/my-scores", methods=["GET"])
@require_manager
def get_my_scores():
    """Return all scores submitted by the current manager for a quarter."""
    session = g.session
    manager_name: str = session["name"]
    is_test: bool = session.get("isTest", False)
    quarter = request.args.get("quarter", "").strip()

    sheets = _sheets(is_test)
    if not quarter:
        settings = sheets.get_settings()
        quarter = settings.get("當前季度") or current_quarter()

    scores = sheets.get_scores_by_manager(quarter, manager_name)
    result = {
        s["empName"]: {
            "scores": s["scores"],
            "special": s["special"],
            "note": s["note"],
            "status": s["status"],
        }
        for s in scores
    }
    return jsonify(result)


# ── GET /api/scoring/all-status ────────────────────────────────────────────

@scoring_bp.route("/all-status", methods=["GET"])
@require_auth
def get_all_status():
    """
    Return scoring progress for all managers (HR / SysAdmin only).
    """
    session = g.session
    if session.get("role") not in ("HR", "系統管理員"):
        return jsonify({"error": "無 HR 權限"}), 403

    is_test: bool = session.get("isTest", False)
    sheets = _sheets(is_test)
    settings = sheets.get_settings()
    quarter = settings.get("當前季度") or current_quarter()

    all_scores = sheets.get_all_scores(quarter)
    score_map: dict[str, dict[str, str]] = {}  # { managerName: { empName: status } }
    for s in all_scores:
        score_map.setdefault(s["managerName"], {})[s["empName"]] = s["status"]

    responsibilities = sheets.get_manager_responsibilities()
    all_employees = sheets.get_all_employees()
    sections_to_employees: dict[str, list[str]] = {}
    for emp in all_employees:
        sections_to_employees.setdefault(emp["section"], []).append(emp["name"])

    # Build per-manager summary
    manager_sections: dict[str, list] = {}
    for r in responsibilities:
        manager_sections.setdefault(r["lineUid"], []).append(r)

    result = []
    accounts = sheets.get_all_accounts()
    manager_accounts = [a for a in accounts if a.get("role") == "主管"]

    for account in manager_accounts:
        uid = account["lineUid"]
        name = account["name"]
        my_resp = manager_sections.get(uid, [])
        my_sections = {r["section"] for r in my_resp}

        emp_names = [
            n for sec in my_sections
            for n in sections_to_employees.get(sec, [])
        ]
        total = len(emp_names)
        manager_scores = score_map.get(name, {})
        scored = sum(1 for n in emp_names if manager_scores.get(n) == "已送出")
        pending = total - scored

        result.append({
            "managerName": name,
            "lineUid": uid,
            "testUid": account.get("testUid", ""),
            "total": total,
            "scored": scored,
            "pending": pending,
        })

    return jsonify(result)


# ── GET /api/scoring/items ─────────────────────────────────────────────────

@scoring_bp.route("/items", methods=["GET"])
@require_auth
def get_score_items():
    """Return the list of scoring items (評分項目)."""
    is_test: bool = g.session.get("isTest", False)
    items = _sheets(is_test).get_score_items()
    return jsonify(items)
