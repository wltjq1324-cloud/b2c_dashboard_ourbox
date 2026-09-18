# b2c_dashboard_ourbox — B2C 매출 대시보드 (Claude Code 프로젝트 지침)

밭(주) 온라인팀(존)의 매출 대시보드. 구글시트(`raw_orders` → `가공_데이터` → `dashboard_cache`) + Apps Script 웹앱(`apps-script-v3.7.js`) + GitHub Pages 정적 페이지. 상세는 `README.md`가 원본이다.

다른 저장소와의 관계·비밀값 위치는 **`batt_ops/docs/REPO-MAP.md`**를 먼저 읽는다. 존과 일하는 규칙은 `batt_ops/CLAUDE.md`와 같다.

## 이 저장소에 없는 것

- **박스(부자재) 실재고·발주 제안은 `promotion_calendar`(`/inven`)에 있다.** 여기 넣지 않는다.
- 아워박스 API 연동 코드 없음. 여기의 `OURBOX_API_URL`은 아워박스가 아니라 우리 Apps Script 웹앱 주소다. API 사실은 `batt_ops/docs/handover/ourbox-api.md`.
- `raw_orders` 시트는 저장소 밖 프로세스로 채워진다(7개 컬럼뿐 — 송장·포장박스 없음).

## 수정 절차 (README 요약)

1. `dashboard.src.html`만 편집한다. **`index.html`은 생성물이라 직접 편집 금지.**
2. `python3 build_index.py` (gzip+base64로 `index.html` 재생성).
3. `dashboard.src.html`을 브라우저로 열어 확인한 뒤 커밋·푸시.
4. Apps Script를 바꿨으면 **배포 관리 → 기존 배포 → 새 버전**으로 올린다. 새 배포를 만들면 URL이 바뀌어 `dashboard.src.html:15`를 고쳐야 한다.

## 구조

- `dashboard.src.html` — 2페이지(매출 현황, 담당자별 성과), Chart.js 4.4.7 CDN, 2단계 로딩(`part=summary` → `part=quality`).
- `apps-script-v3.7.js` — `doGet(mode=cache&part=…)`, `dashboard_cache` 시트 + CacheService 6h. 갱신은 스프레드시트 메뉴 "📊 아워박스"에서 수동.
- `edit.html` — 별개 앱(프로모션 캘린더 편집, 다른 Apps Script). `promotion_calendar` 저장소의 캘린더와 혼동 주의.
