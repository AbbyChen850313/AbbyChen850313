"""
Flask application factory.
"""

from __future__ import annotations

import logging
import os

from flask import Flask, jsonify
from flask_cors import CORS

import config


def create_app() -> Flask:
    app = Flask(__name__)

    # ── Logging ────────────────────────────────────────────────────────────
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    # ── CORS ───────────────────────────────────────────────────────────────
    CORS(
        app,
        resources={r"/api/*": {"origins": config.ALLOWED_ORIGINS}},
        supports_credentials=True,
    )

    # ── Blueprints ─────────────────────────────────────────────────────────
    from routes.auth import auth_bp
    from routes.dashboard import dashboard_bp
    from routes.scoring import scoring_bp
    from routes.admin import admin_bp

    app.register_blueprint(auth_bp, url_prefix="/api/auth")
    app.register_blueprint(dashboard_bp, url_prefix="/api")
    app.register_blueprint(scoring_bp, url_prefix="/api/scoring")
    app.register_blueprint(admin_bp, url_prefix="/api/admin")

    # ── Health check ───────────────────────────────────────────────────────
    @app.route("/health")
    def health():
        return jsonify({"status": "ok"})

    # ── Global error handlers ──────────────────────────────────────────────
    @app.errorhandler(404)
    def not_found(e):
        return jsonify({"error": "找不到此端點"}), 404

    @app.errorhandler(405)
    def method_not_allowed(e):
        return jsonify({"error": "不支援此 HTTP 方法"}), 405

    @app.errorhandler(500)
    def internal_error(e):
        logging.getLogger(__name__).exception("Unhandled error")
        return jsonify({"error": "伺服器內部錯誤，請稍後再試"}), 500

    return app


if __name__ == "__main__":
    app = create_app()
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
