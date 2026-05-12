# B2C Ourbox Sales Dashboard

아워박스 B2C 매출 진척도 대시보드와 Google Sheets 기반 데이터 파이프라인 작업 기록입니다.

## 현재 운영 구조

- 운영 대시보드: `index.html`
- 캐시 테스트 대시보드: `sales-progress-dashboard-prototype-20260512-cache.html`
- 기존 프로토타입: `sales-progress-dashboard-prototype-20260512.html`
- 데이터 원본: Google Sheets `raw_orders`, `가공_데이터`, `map_channel`, `map_product`, `map_cost`
- Apps Script 현재 기준: v3.4 cache patch

운영 `index.html`은 보호 대상입니다. 새 실험은 별도 HTML 파일에서 먼저 검증한 뒤 반영합니다.

## 2026-05-12 작업 요약

### Google Sheets / Apps Script

- 기존 v3.3 로직을 v3.4로 확장했습니다.
- `refreshProcessedData()` 실행 시 `가공_데이터`를 갱신한 뒤 `dashboard_cache` 숨김 시트를 자동 생성/갱신합니다.
- `dashboard_cache`는 사람이 직접 수정하는 시트가 아니라 대시보드 로딩 속도를 줄이기 위한 중간 캐시입니다.
- 기존 대시보드 호환을 위해 기본 `doGet()` 응답은 기존처럼 `orders` 전체 응답을 유지합니다.
- 새 대시보드는 `?mode=cache` 요청에서만 `dashboardCache` 요약 응답을 받습니다.

### dashboard_cache 구성

- `baseRows`: 일자 × 채널 × 상품군 × 담당자 단위 요약
- `productRows`: 일자 × 채널 × 상품군 × 담당자 × 상품 단위 요약
- `qualityRows`: 미매핑, 원가 0/1, 매출 0, 마진 음수 등 품질 이슈 행만 분리

최근 실행 결과:

- 처리 행 수: `52,713`
- 고유 주문 수: `44,646`
- 캐시 행 수: `base 1,695 / product 3,603 / quality 6,064`

### HTML 대시보드

- 캐시 테스트 페이지를 추가했습니다.
- URL: https://wltjq1324-cloud.github.io/b2c_dashboard_ourbox/sales-progress-dashboard-prototype-20260512-cache.html
- 새 HTML은 Apps Script에 `mode=cache`를 붙여 요청합니다.
- 캐시 응답이 오면 `dashboardCache.baseRows`, `productRows`, `qualityRows`를 우선 사용합니다.
- 캐시 응답이 없을 때는 기존 `orders` 응답을 브라우저에서 요약하는 폴백을 유지합니다.

### 성능 확인

- 기존 체감 로딩: 약 `15초`
- 캐시 적용 후 체감 로딩: 약 `7초`
- 약 절반 수준으로 개선되었습니다.

완전한 즉시 로딩까지는 아직 아닙니다. Google Apps Script 웹앱 호출 자체의 왕복 시간, Google Sheets 읽기 시간, `qualityRows` 전송량이 남아 있습니다.

## 40만 행 증가 시 예상

기존 3.3 방식처럼 40만 행 전체를 JSON으로 내려주면 1분 이상 걸리거나 timeout 가능성이 큽니다.

현재 v3.4 캐시 방식은 원본 전체 행이 아니라 `dashboard_cache` 크기에 비례합니다. 현재 패턴이 유지된다면 40만 행에서도 기존 방식보다는 훨씬 안전하지만, `qualityRows`가 같이 커지면 로딩은 다시 늘 수 있습니다.

대략 예상:

- 현재 패턴 유지: `10~25초`
- 품질 이슈 행 증가: `30초 이상`
- 기존 raw 전체 전송 방식: `1~2분 이상` 또는 timeout 위험

## 다음 최적화 포인트

### 1. qualityRows 지연 로딩

현재 새 대시보드는 상세 품질 확인을 위해 `qualityRows`까지 첫 응답에 함께 받습니다. 이 데이터가 커지면 캐시 방식에서도 로딩 병목이 됩니다.

다음 구조가 권장됩니다.

- 첫 로딩: `baseRows`, `productRows`, 품질 이슈별 요약 숫자만 전송
- 사용자가 품질 항목을 클릭할 때: 해당 이슈 상세 행만 별도 요청
- 예상 효과: 40만 행 환경에서 첫 화면 로딩 안정성 크게 개선

### 2. productRows 범위 축소 또는 별도 요청

상품 TOP 15용으로 전체 `productRows`를 매번 받는 구조입니다. 상품 수가 늘면 이 영역도 커집니다.

개선 방향:

- 첫 응답에는 기간별/채널별 TOP N만 포함
- 상품 상세 테이블은 필터 변경 시 별도 요청
- 또는 Apps Script에서 최근 조회 조건별 product cache를 별도 생성

### 3. 캐시 시트 분리

현재 `dashboard_cache` 한 시트에 `meta`, `baseRows`, `productRows`, `qualityRows` JSON chunk를 모두 저장합니다.

데이터가 커지면 아래처럼 분리하는 것이 안정적입니다.

- `dashboard_cache_meta`
- `dashboard_cache_base`
- `dashboard_cache_product`
- `dashboard_cache_quality_summary`
- `dashboard_cache_quality_detail`

### 4. 품질 상세 전용 endpoint

Apps Script `doGet(e)`에 아래 mode를 추가하는 방향이 좋습니다.

- `mode=cache`: 첫 화면 요약
- `mode=quality&issue=channelUnmapped`: 특정 품질 상세
- `mode=orders`: 기존 전체 orders 디버그용

### 5. 운영 반영 전 검증 루틴

새 실험 파일에서 먼저 확인합니다.

1. `refreshProcessedData()` 실행
2. `dashboard_cache` 갱신 로그 확인
3. 캐시 테스트 대시보드 로딩 확인
4. 기존 운영 대시보드가 깨지지 않는지 확인
5. 수치 검산: 매출, 주문 수, 마진, 담당자별 집계
6. 문제가 없을 때만 운영 `index.html` 반영 검토

## 주의 사항

- `dashboard_cache`는 삭제하거나 직접 수정하지 않습니다.
- 삭제해도 `refreshProcessedData()`를 다시 실행하면 재생성되지만, 새 대시보드가 느린 폴백을 탈 수 있습니다.
- Apps Script는 저장만으로 웹앱 `/exec`에 반영되지 않습니다.
- 웹앱 응답이 바뀌어야 할 때는 Apps Script에서 기존 웹앱 배포를 `새 버전`으로 갱신해야 합니다.
- 기존 배포를 수정해야 URL이 유지됩니다. 새 배포를 만들면 HTML의 `API_URL`도 바꿔야 합니다.

## 주요 커밋

- `62ac7b7e80383ed78ccc01b253f473f6dfb18eef`: 캐시 기반 테스트 HTML 추가
- `79ee8f51e88652f39602addc4bc5f61c3decdec2`: Apps Script 캐시 패치 스니펫 추가
