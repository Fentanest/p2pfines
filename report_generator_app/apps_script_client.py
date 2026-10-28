from __future__ import annotations

from pathlib import Path
from typing import Any

from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = [
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.deployments",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


class AppsScriptInstallError(RuntimeError):
    pass


class AppsScriptClient:
    def __init__(self, service_account_json: str | Path | None = None, credentials: Any | None = None):
        if credentials is not None:
            creds = credentials
        elif service_account_json:
            creds = ServiceAccountCredentials.from_service_account_file(str(service_account_json), scopes=SCOPES)
        else:
            raise AppsScriptInstallError("Apps Script 설치용 인증 정보가 없습니다.")
        self.service = build("script", "v1", credentials=creds, cache_discovery=False)

    def create_bound_project(self, spreadsheet_id: str, title: str) -> str:
        try:
            result = self.service.projects().create(
                body={"title": title, "parentId": spreadsheet_id}
            ).execute()
            return result["scriptId"]
        except HttpError as exc:
            raise AppsScriptInstallError(_format_script_error(exc)) from exc

    def update_content(self, script_id: str, code: str) -> None:
        body: dict[str, Any] = {
            "files": [
                {
                    "name": "Code",
                    "type": "SERVER_JS",
                    "source": code,
                },
                {
                    "name": "appsscript",
                    "type": "JSON",
                    "source": """{
  "timeZone": "Asia/Seoul",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "executionApi": {
    "access": "MYSELF"
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly"
  ]
}""",
                },
            ]
        }
        try:
            self.service.projects().updateContent(scriptId=script_id, body=body).execute()
        except HttpError as exc:
            raise AppsScriptInstallError(_format_script_error(exc)) from exc


def _format_script_error(exc: HttpError) -> str:
    text = str(exc)
    if "script.googleapis.com" in text and "disabled" in text.lower():
        return (
            "Apps Script API가 Google Cloud 프로젝트에서 비활성화되어 있어 직접 설치할 수 없습니다. "
            "Google Cloud Console에서 script.googleapis.com을 활성화한 뒤 다시 시도해주세요."
        )
    if "User has not enabled the Apps Script API" in text or "script.google.com/home/usersettings" in text:
        return (
            "Apps Script API의 사용자 설정이 꺼져 있어 직접 설치할 수 없습니다. "
            "https://script.google.com/home/usersettings 에서 Apps Script API를 켠 뒤 다시 시도해주세요. "
            "서비스 계정 JSON만 사용하는 현재 방식에서는 이 사용자 설정을 제가 API로 대신 켤 수 없습니다."
        )
    if "PERMISSION_DENIED" in text or "Permission" in text:
        return (
            "Apps Script 직접 설치 권한이 없습니다. OAuth로 승인한 Google 계정이 해당 스프레드시트의 편집 권한을 "
            "가졌는지 확인해주세요. 상세: " + text[:500]
        )
    return "Apps Script API 오류: " + text[:800]
