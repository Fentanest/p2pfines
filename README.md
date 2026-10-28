# Jjukfund 보고서 온라인 업로더

이 프로젝트는 로컬에서 보고서 값을 산출하지 않는다. 로컬 웹 UI는 백데이터 파일을 읽어 선택한 원천 컬럼만 구글 스프레드시트에 업로드하고, 보고서 계산/검증/추적은 구글 스프레드시트의 Apps Script에서 수행한다.

## 실행

```powershell
& 'C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' app.py
```

웹 UI: http://127.0.0.1:8787

### Windows 런처

소스 상태에서 경량 데스크톱 런처를 실행하려면:

```powershell
python desktop_launcher.py
```

`웹서버 시작`을 누르면 로컬 서버가 시작되고 기본 브라우저에서 웹 UI가 열린다. 실행 중에는 버튼이 `종료`로 바뀌며, 버튼을 누르거나 런처 창을 닫으면 서버도 함께 종료된다.

EXE 빌드:

```powershell
python -m pip install -r requirements-build.txt
.\build_windows.ps1 -Python python
```

결과물은 `dist\JjukfundLauncher\JjukfundLauncher.exe`에 생성된다. `onedir` 배포물이므로 `JjukfundLauncher` 폴더 전체를 함께 배포해야 한다.

EXE에서 생성되는 설정, 인증 파일, 임시 업로드, 결과 파일은 `%LOCALAPPDATA%\Jjukfund`에 보존된다. 서비스 계정 및 OAuth JSON은 EXE에 포함되지 않으며 기존 웹 UI에서 업로드한다.

## 파일 이름 매핑

| 우리 이름 | 구글시트 원천 탭 |
|---|---|
| 전체계약 | 계약명세 |
| 기준일투자자 | 기준일투자자 |
| 기준일계약 | 기준일계약명세 |

## 암호 파일

기본 엑셀 비밀번호는 `118050`으로 저장되어 있다. 파일 추가 시 암호화가 감지되면 UI가 비밀번호를 묻고 같은 파일을 다시 파싱한다.

## 온라인 계산 설치

1. 웹 UI에서 `구글 문서 탭 초기화`를 실행한다.
2. 웹 UI에서 `Google OAuth 인증`을 완료한다.
3. `Apps Script 직접 설치 시도`를 눌러 각 문서의 Apps Script 프로젝트에 코드를 설치한다.
4. 이후 구글시트의 `실행` 탭 또는 메뉴 `Jjukfund > 보고서 재계산`으로 온라인 산출을 실행한다.

## 추적 구조

원천 탭 → `계산_계약명세`, `계산_기준일투자자`, `계산_기준일계약명세` → 보고서 탭 → `검증`, `실행로그`
