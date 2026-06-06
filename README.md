# ⚾ KBO 카드 컬렉션 (KBO SCC Card Collection)

7-ELEVEN SCC KBO 야구 카드 보유 체크리스트 PWA.

- **연도(2025·2026) › 구단 › 선수 목록** 구조
- 체크박스 탭으로 보유 표시, 진행률 자동 계산
- 기록은 브라우저(기기)에 자동 저장 (localStorage)
- iOS: Safari → 공유 → **홈 화면에 추가** 하면 앱처럼 사용 (오프라인 동작)

## 사용

GitHub Pages로 배포되어 있습니다. 휴대폰 브라우저로 접속 후 홈 화면에 추가하세요.

## 구성

- `index.html` — 앱 + 데이터 내장 (단일 파일)
- `manifest.json`, `sw.js` — PWA 매니페스트 / 서비스 워커(오프라인 캐시)
- `icon-*.png` — 앱 아이콘

데이터 출처: SCC 2025 / 2026 KBO 컬렉션 카드 체크리스트.
