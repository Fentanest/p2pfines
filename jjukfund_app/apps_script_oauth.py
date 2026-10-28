from __future__ import annotations

import json
import secrets
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request as UrlRequest
from urllib.request import urlopen

from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials

from .apps_script_client import SCOPES
from .config import CREDENTIAL_DIR
from .paths import DATA_ROOT


TOKEN_PATH = CREDENTIAL_DIR / "apps-script-oauth-token.json"
STATE_PATH = CREDENTIAL_DIR / "apps-script-oauth-state.json"
DEFAULT_REDIRECT_URI = "http://127.0.0.1:8787/oauth2callback"


class AppsScriptOAuthError(RuntimeError):
    pass


def find_oauth_client_json(preferred_path: str | None = None) -> Path:
    candidates: list[Path] = []
    if preferred_path:
        candidates.append(Path(preferred_path))
    candidates.extend(sorted(CREDENTIAL_DIR.glob("client_secret*.json")))
    candidates.extend(sorted(CREDENTIAL_DIR.glob("*oauth*.json")))
    candidates.extend(sorted(DATA_ROOT.glob("client_secret*.json")))
    candidates.extend(sorted(DATA_ROOT.glob("*oauth*.json")))
    candidates.extend(sorted(DATA_ROOT.glob("*.json")))

    seen: set[Path] = set()
    for path in candidates:
        path = path.resolve()
        if path in seen or not path.exists():
            continue
        seen.add(path)
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if _client_section(data):
            return path
    raise AppsScriptOAuthError("Apps Script OAuth 클라이언트 JSON을 찾지 못했습니다.")


def oauth_status(preferred_path: str | None = None) -> dict[str, Any]:
    status: dict[str, Any] = {
        "hasClient": False,
        "clientJson": "",
        "hasToken": TOKEN_PATH.exists(),
        "valid": False,
    }
    try:
        client_path = find_oauth_client_json(preferred_path)
        status["hasClient"] = True
        status["clientJson"] = str(client_path)
    except AppsScriptOAuthError as exc:
        status["error"] = str(exc)
        return status

    try:
        creds = load_user_credentials(refresh=True)
        status["valid"] = bool(creds.valid)
        status["expired"] = bool(creds.expired)
        status["scopes"] = list(creds.scopes or [])
    except AppsScriptOAuthError as exc:
        status["error"] = str(exc)
    return status


def build_authorization_url(
    preferred_path: str | None = None,
    redirect_uri: str = DEFAULT_REDIRECT_URI,
) -> dict[str, str]:
    client_path = find_oauth_client_json(preferred_path)
    client = _load_client(client_path)
    redirects = client.get("redirect_uris") or []
    if redirect_uri not in redirects:
        raise AppsScriptOAuthError(
            f"OAuth JSON에 리디렉션 URI가 없습니다: {redirect_uri}"
        )

    state = secrets.token_urlsafe(32)
    STATE_PATH.write_text(
        json.dumps(
            {
                "state": state,
                "redirect_uri": redirect_uri,
                "client_json": str(client_path),
                "created_at": time.time(),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    query = urlencode(
        {
            "client_id": client["client_id"],
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": " ".join(SCOPES),
            "access_type": "offline",
            "include_granted_scopes": "true",
            "prompt": "consent",
            "state": state,
        }
    )
    return {
        "authUrl": f"{client.get('auth_uri', 'https://accounts.google.com/o/oauth2/auth')}?{query}",
        "clientJson": str(client_path),
        "redirectUri": redirect_uri,
    }


def finish_oauth_callback(code: str, state: str) -> dict[str, Any]:
    if not code:
        raise AppsScriptOAuthError("Google 인증 코드가 비어 있습니다.")
    stored = _load_state()
    if not state or state != stored.get("state"):
        raise AppsScriptOAuthError("OAuth state 검증에 실패했습니다. 인증을 다시 시작하세요.")
    if time.time() - float(stored.get("created_at") or 0) > 15 * 60:
        raise AppsScriptOAuthError("OAuth 인증 요청이 만료되었습니다. 인증을 다시 시작하세요.")

    client_path = Path(stored["client_json"])
    client = _load_client(client_path)
    token_payload = _exchange_code_for_token(
        client=client,
        code=code,
        redirect_uri=stored["redirect_uri"],
    )
    _save_token(client, token_payload)
    STATE_PATH.unlink(missing_ok=True)
    return {"ok": True, "clientJson": str(client_path), "hasRefreshToken": bool(token_payload.get("refresh_token"))}


def load_user_credentials(refresh: bool = True) -> Credentials:
    if not TOKEN_PATH.exists():
        raise AppsScriptOAuthError("Google OAuth 인증이 아직 완료되지 않았습니다.")
    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), scopes=SCOPES)
    if refresh and creds.expired and creds.refresh_token:
        creds.refresh(GoogleAuthRequest())
        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    if not creds.valid:
        raise AppsScriptOAuthError("저장된 Google OAuth 토큰이 유효하지 않습니다. 다시 인증하세요.")
    return creds


def _load_client(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    client = _client_section(data)
    if not client:
        raise AppsScriptOAuthError(f"OAuth 클라이언트 JSON 형식이 아닙니다: {path}")
    if not client.get("client_id") or not client.get("client_secret"):
        raise AppsScriptOAuthError(f"OAuth JSON에 client_id/client_secret이 없습니다: {path}")
    return client


def _client_section(data: dict[str, Any]) -> dict[str, Any] | None:
    web = data.get("web")
    if isinstance(web, dict):
        return web
    installed = data.get("installed")
    if isinstance(installed, dict):
        return installed
    return None


def _load_state() -> dict[str, Any]:
    if not STATE_PATH.exists():
        raise AppsScriptOAuthError("진행 중인 OAuth 인증 요청이 없습니다.")
    return json.loads(STATE_PATH.read_text(encoding="utf-8"))


def _exchange_code_for_token(client: dict[str, Any], code: str, redirect_uri: str) -> dict[str, Any]:
    body = urlencode(
        {
            "code": code,
            "client_id": client["client_id"],
            "client_secret": client["client_secret"],
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
        }
    ).encode("utf-8")
    request = UrlRequest(
        client.get("token_uri", "https://oauth2.googleapis.com/token"),
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise AppsScriptOAuthError(f"Google OAuth 토큰 교환 실패: {detail}") from exc
    except URLError as exc:
        raise AppsScriptOAuthError(f"Google OAuth 요청 실패: {exc}") from exc
    if "error" in payload:
        raise AppsScriptOAuthError(f"Google OAuth 토큰 교환 실패: {payload}")
    return payload


def _save_token(client: dict[str, Any], token_payload: dict[str, Any]) -> None:
    refresh_token = token_payload.get("refresh_token")
    if not refresh_token and TOKEN_PATH.exists():
        try:
            previous = json.loads(TOKEN_PATH.read_text(encoding="utf-8"))
            refresh_token = previous.get("refresh_token")
        except Exception:
            refresh_token = None

    expiry = datetime.now(timezone.utc) + timedelta(seconds=int(token_payload.get("expires_in") or 3600))
    creds = Credentials(
        token=token_payload.get("access_token"),
        refresh_token=refresh_token,
        token_uri=client.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=client["client_id"],
        client_secret=client["client_secret"],
        scopes=SCOPES,
        expiry=expiry,
    )
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
