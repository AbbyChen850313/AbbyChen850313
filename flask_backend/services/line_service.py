"""
LINE API operations: token verification, profile fetch, push messages.
"""

from __future__ import annotations

import logging

import requests

import config

logger = logging.getLogger(__name__)

_LINE_VERIFY_URL = "https://api.line.me/oauth2/v2.1/verify"
_LINE_PROFILE_URL = "https://api.line.me/v2/profile"
_LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"


def verify_access_token(access_token: str) -> dict | None:
    """
    Verify a LIFF access token and return the user profile dict,
    or None if verification fails.

    Returns: { "userId": str, "displayName": str, "pictureUrl": str }
    """
    verify_resp = requests.get(
        _LINE_VERIFY_URL,
        params={"access_token": access_token},
        timeout=10,
    )
    if verify_resp.status_code != 200:
        logger.warning("LINE token verification failed: %s", verify_resp.text)
        return None

    profile_resp = requests.get(
        _LINE_PROFILE_URL,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=10,
    )
    if profile_resp.status_code != 200:
        logger.warning("LINE profile fetch failed: %s", profile_resp.text)
        return None

    return profile_resp.json()


def push_message(line_uid: str, text: str, is_test: bool = False) -> bool:
    """Send a LINE push text message. Returns True on success."""
    token = config.line_channel_token(is_test=is_test)
    resp = requests.post(
        _LINE_PUSH_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "to": line_uid,
            "messages": [{"type": "text", "text": text}],
        },
        timeout=10,
    )
    if resp.status_code != 200:
        logger.error("LINE push failed (uid=%s): %s", line_uid, resp.text)
        return False
    return True
