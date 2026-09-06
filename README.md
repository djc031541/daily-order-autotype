# 정비통제업무일지 초안 생성기

Daily Work Order 엑셀을 올리면 **정비통제업무일지 docx 초안**을 만들어 줍니다.
`항공기 주요 작업 사항 (정비 1)` / `(정비 2)` 칸에 작업 제목이 자동으로 채워지고,
작업 결과는 근무가 끝난 뒤 `=>` 뒤에 직접 입력합니다.

## 배포 (Streamlit Community Cloud)

이 저장소를 그대로 배포하면 사내 어느 PC에서든 브라우저만으로 쓸 수 있습니다.
설치할 것이 없어 공용 PC에서도 동작합니다.

1. https://share.streamlit.io 에 GitHub 계정으로 로그인
2. **Create app → Deploy a public app from GitHub** 선택
   - Repository: `djc031541/daily-order-autotype`
   - Branch: `main`
   - Main file path: `streamlit_app.py`
   - 이 저장소는 **private** 이므로 Streamlit 에 저장소 접근 권한을 한 번 승인해야 합니다
3. 배포 후 **Settings → Sharing** 에서 **"Only specific people can view this app"** 로 바꾸고
   볼 사람 이메일만 등록 — **반드시 하세요.** 기본값은 링크만 알면 누구나 열립니다.

`requirements.txt` 와 `.python-version` 이 있어 그 외 설정은 필요 없습니다.

### 올리는 데이터에 대해

Daily Work Order 에는 근무자 실명·사내 연락처·항공기 결함 내용이 들어 있습니다.
Community Cloud 에 올리면 그 파일이 **Streamlit(Snowflake) 서버에서 처리**됩니다.
파일을 디스크에 남기지는 않지만(처리 후 임시파일만 사용), 사내 데이터를 외부 호스팅으로
보내는 것 자체가 맞는지는 한 번 확인해 보시는 게 좋습니다.
사내에서만 쓰려면 집 맥북 + Tailscale Funnel 구성(TRINITY LAN 서버와 같은 방식)이 대안입니다.

## 로컬 실행

Finder에서 `run.command` 더블클릭. 또는:

```bash
.venv/bin/streamlit run streamlit_app.py
```

브라우저에서 Daily Work Order 엑셀을 올리고, 근무(날짜 + D/N)를 고른 뒤 docx를 내려받습니다.

명령줄:

```bash
.venv/bin/python cli.py "Daily Work Order (2026-09-02).xlsx"              # 파일에 든 근무 목록
.venv/bin/python cli.py "Daily Work Order (2026-09-02).xlsx" --shift N    # 야간 초안 생성
.venv/bin/python cli.py "Daily Work Order (2026-09-02).xlsx" --shift D --print   # 본문만 출력
```

생성물은 `output/정비통제업무일지_(03-SEP-26)D.docx` 형식으로 저장됩니다.

## 무엇이 채워지고 무엇이 안 채워지나

| 칸 | 처리 |
|---|---|
| `DATE : 2026. 09. 03. (목) D` | 자동 |
| 항공기 주요 작업 사항 **(정비 1)** | 자동 — Daily Work Order 의 **GMP** 작업 |
| 항공기 주요 작업 사항 **(정비 2)** | 자동 — Daily Work Order 의 **ICN** 작업 |
| 항공기 주요 작업 사항 (국내. 해외) | **비워 둠** (IRR 사항이라 초안 대상 아님) |
| 인수인계자 / 기상 / FLT PLAN / 근무 인원 | **비워 둠** (직접 입력) |
| 항공기 IRR / SKD MAINT / SEAT BLOCK | **비워 둠** (직접 입력) |

지난 일지 값이 그대로 남아 잘못 나가는 일이 없도록, 초안을 만들 때 위 칸들은 항상 비웁니다.
(그대로 두고 싶으면 CLI `--keep-previous`, 앱 사이드바에서 체크 해제)

## 출력 형식

기종별로 묶고, 기번별로 번호를 매깁니다. 순서는 엑셀에 나온 순서 그대로입니다.

```
[B738]
HL8547 1) CL-BI-WEEKLY CHECK=>
         2) AD-ULTRASONIC AND EXTERNAL DETAILED INSPECTION - LH/RH STA 663.75 ...=>
[B38M]
HL8751 1) TRP-REPLACEMENT - AIRCRAFT BATTERY (P/N: BA35-01, ...)=>
```

- `Type-작업제목` 형태 (`CL-`, `AD-`, `NRC-`, `EO-`, `SEM-`, `TRP-`, `ETC-`)
- Type 이 `DEF` 이면 `DEF-제목`, 엑셀 제목이 `[DEF CLR] …` 이면 `DEF CLR: 제목`
- 기번은 `HL` 이 없으면 붙여 줍니다 (`8547` → `HL8547`)
- 서식(대명체 8pt Bold)은 템플릿의 기존 서식을 그대로 복제합니다

## 엑셀 읽는 방식

한 파일에 `Date … Shift` 행으로 시작하는 블록이 여러 개 들어 있고
(예: `2026-09-02 N`, `2026-09-03 D`), GMP / ICN / 지방 블록이 따로 쌓여 있습니다.
프로그램은 같은 (날짜, Shift) 블록을 합친 뒤 `STA` 열로 정비 1 / 정비 2 를 나눕니다.

컬럼은 위치가 아니라 **헤더 텍스트**(`REG' NO.`, `Type`, `Order Description / Title / Action` …)로
찾습니다. 엑셀 양식에서 열 순서가 바뀌어도 헤더만 같으면 동작합니다.

주의해서 볼 만한 처리:

- 같은 항공기에서 `Type` 칸이 비어 있는 연속 행은 **바로 위 행의 Type** 을 이어받습니다.
- `Order Description` 이 비어 있는 행은 일지에 올릴 내용이 없다고 보고 건너뜁니다
  (도착/출발 시각만 있는 GRD 항공기 등).
- 기번·기종·Type 이 비어 있어 사람이 확인해야 하는 행은 실행 후 **확인 필요** 목록에 표시됩니다.

## 템플릿 교체

`template/정비통제업무일지_TEMPLATE.docx` 가 서식 원본입니다. 서식만 있고 내용은 비어 있습니다.
양식이 바뀌면 최신 업무일지 docx 를 이 이름으로 덮어쓰거나, 앱 사이드바에서 올리면 됩니다.

> **저장소에 새 템플릿을 올릴 때는 반드시 먼저 지우고 올리세요.**
> 업무일지 원본에는 인수인계자 실명·IRR 결함 내용이 파일 안에 그대로 남아 있습니다.
> ```bash
> .venv/bin/python scripts/scrub_template.py template/정비통제업무일지_TEMPLATE.docx
> ```
> 이 스크립트가 인수인계자·기상·근무인원·IRR·SEAT BLOCK·작업사항과 문서 작성자 메타데이터까지 비웁니다.
프로그램은 표 안의 제목 행(`항공기 주요 작업 사항 (정비 1)` 등)을 **텍스트로 찾아서**
그 다음 행에 내용을 넣기 때문에, 행이 몇 개 늘거나 줄어도 따라갑니다.

## 구성

```
streamlit_app.py          Streamlit 업로드 UI (배포 진입점)
cli.py                    명령줄
worksheet_auto/parser.py       Daily Work Order 엑셀 파싱
worksheet_auto/formatter.py    업무일지 본문 텍스트 생성
worksheet_auto/docx_writer.py  템플릿 채우기 (서식 유지)
template/                 업무일지 서식 원본
scripts/scrub_template.py 템플릿에서 실데이터 지우기
samples/                  예시 입력 — 사내 실데이터라 저장소에서 제외(.gitignore)
output/                   생성물
```
