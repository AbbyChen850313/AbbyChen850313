"""
/api/admin  — HR and SysAdmin management operations.
"""

from __future__ import annotations

import logging

from flask import Blueprint, g, jsonify, request

from services.auth_service import require_hr, require_sysadmin
from services.sheets_service import SheetsService

logger = logging.getLogger(__name__)
admin_bp = Blueprint("admin", __name__)


def _sheets(is_test: bool) -> SheetsService:
    return SheetsService(is_test=is_test)


# ── GET /api/admin/settings ────────────────────────────────────────────────

@admin_bp.route("/settings", methods=["GET"])
@require_hr
def get_settings():
    is_test: bool = g.session.get("isTest", False)
    return jsonify(_sheets(is_test).get_settings())


# ── POST /api/admin/settings ───────────────────────────────────────────────

@admin_bp.route("/settings", methods=["POST"])
@require_hr
def update_settings():
    is_test: bool = g.session.get("isTest", False)
    body = request.get_json(silent=True) or {}
    if not body:
        return jsonify({"error": "沒有要更新的設定"}), 400
    _sheets(is_test).update_settings(body)
    return jsonify({"success": True})


# ── GET /api/admin/employees ───────────────────────────────────────────────

@admin_bp.route("/employees", methods=["GET"])
@require_hr
def get_employees():
    is_test: bool = g.session.get("isTest", False)
    return jsonify(_sheets(is_test).get_all_employees())


# ── POST /api/admin/employees/sync ────────────────────────────────────────

@admin_bp.route("/employees/sync", methods=["POST"])
@require_hr
def sync_employees():
    """Sync employee list from HR spreadsheet."""
    is_test: bool = g.session.get("isTest", False)
    count = _sheets(is_test).sync_employees_from_hr()
    return jsonify({"success": True, "count": count})


# ── GET /api/admin/score-items ─────────────────────────────────────────────

@admin_bp.route("/score-items", methods=["GET"])
@require_hr
def get_score_items():
    is_test: bool = g.session.get("isTest", False)
    return jsonify(_sheets(is_test).get_score_items())


# ── GET /api/admin/responsibilities ───────────────────────────────────────

@admin_bp.route("/responsibilities", methods=["GET"])
@require_hr
def get_responsibilities():
    """Return the manager-section weight table (主管權重)."""
    is_test: bool = g.session.get("isTest", False)
    return jsonify(_sheets(is_test).get_manager_responsibilities())


# ── POST /api/admin/refresh-roles ─────────────────────────────────────────

@admin_bp.route("/refresh-roles", methods=["POST"])
@require_sysadmin
def refresh_roles():
    """Re-derive roles for all accounts from HR Sheet. SysAdmin only."""
    # This is a complex operation that reads HR Sheet job titles and reclassifies
    # roles. For now return not-implemented and let the user do it via Sheets.
    return jsonify({"error": "此功能尚在開發中，請直接修改 LINE帳號 表的角色欄位"}), 501
