# INTAX Dashboard — Claude Code 작업 가이드

## 프로젝트 개요
세무회계 컨설팅 법인 **INTAX**의 내부 경영 대시보드.
단일 HTML 파일(`index.html`, ~90KB) 로 구성된 클라이언트 사이드 앱.

- **GitHub**: `iukimblock-collab/intax-dashboard` (main 브랜치)
- **언어**: HTML + CSS + Vanilla JS (빌드 도구 없음)
- **외부 라이브러리**: Chart.js 4.4.0 (CDN), Noto Sans KR / IBM Plex Mono (Google Fonts)
- **데이터 저장**: `localStorage` (서버 없음)

## 파일 구조
```
index.html      # 전체 앱 (CSS + HTML + JS 단일 파일)
.env            # GitHub API 키 (GITHUB_API_KEY)
github_client.py  # API 키 로드 예시 (Python)
.gitignore      # .env 제외 설정
```

## 디자인 시스템

### CSS 변수 (`:root`)
| 변수 | 용도 |
|------|------|
| `--bg` `--surface` `--surface2` `--surface3` | 배경 레이어 |
| `--text` `--muted` `--dim` | 텍스트 계층 |
| `--accent` / `--accent-a` | 파란 강조색 |
| `--green` `--orange` `--yellow` `--red` | 상태색 |
| `--font` `--mono` | Noto Sans KR / IBM Plex Mono |
| `--sw` `--hh` | 사이드바 너비(224px) / 헤더 높이(54px) |

### 공통 컴포넌트 클래스
- `.card` — 데이터 카드
- `.g2` `.g3` `.g4` — 2/3/4열 그리드
- `.btn` `.btn-primary` `.btn-ghost` `.btn-danger` — 버튼
- `.tag` `.tag-green` `.tag-blue` 등 — 상태 배지
- `.tbl-wrap` + `table` — 테이블
- `.modal` + `#modal-overlay` — 모달
- `.tabs` + `.tab` — 탭 UI
- `.filter-bar` — 필터 영역
- `.page` / `.page.active` — 페이지 전환 (fadeIn 애니메이션)

## 앱 구조

### 인증 시스템
- `#login-screen` — 로그인/계정 생성 화면 (z-index: 2000)
- `doLogin()` / `doRegister()` / `doLogout()` 함수
- 계정 정보는 `localStorage`에 저장
- 기본 대표 계정: `김인욱` / `12345678`

### 역할(Role) 시스템
- `role-boss` (대표) / 직원 두 가지 역할
- `body.role-boss .role-boss-only` — 대표만 보이는 UI
- `body:not(.role-boss) .btn-edit-only` — 직원은 편집 불가 (opacity 0.3)

### 페이지 구성 (`data-page` 속성으로 전환)
| 페이지 ID | 메뉴명 | 접근 |
|-----------|--------|------|
| `dashboard` | 대시보드 | 전체 |
| `clients` | 거래처 현황 | 전체 |
| `fees` | 기장료 관리 | 대표만 |
| `notices` | 안내문 발송 | 전체 |
| `staff` | 직원 현황 | 전체 |
| `settings` | 설정 / 권한 | 전체 |

### 레이아웃
```
#shell (flex)
├── #sidebar (224px)
│   ├── #logo
│   ├── #nav (.nav-item[data-page])
│   ├── #sidebar-foot (사용자 정보)
│   └── #logout-btn
└── #main (flex-1)
    ├── #topbar (54px) — 제목, 검색, 액션 버튼
    └── #page-view — .page 들 (active만 표시)
```

### 반응형
- `≤960px`: 사이드바 축소 (60px, 아이콘만)
- `≤640px`: 사이드바 숨김, 그리드 단순화

## 작업 시 주의사항

1. **단일 파일 구조** — CSS, HTML, JS 모두 `index.html` 안에 있음. 분리하지 말 것.
2. **빌드 없음** — npm/번들러 없이 브라우저에서 바로 열어 테스트.
3. **localStorage 의존** — 데이터 초기화 시 브라우저 저장소 삭제 필요.
4. **Chart.js** — CDN 로드, 차트 인스턴스는 업데이트 전 `destroy()` 필요.
5. **CSS 변수 우선** — 색상/간격 하드코딩 금지, 항상 변수 사용.
6. **한국어 UI** — 모든 사용자 표시 텍스트는 한국어 유지.

## GitHub 연동
```powershell
# API 키 로드 후 레포 내용 조회 예시
$key = (Get-Content .env | Where-Object { $_ -match "GITHUB_API_KEY" }) -replace "GITHUB_API_KEY=", ""
$headers = @{ Authorization = "token $key"; Accept = "application/vnd.github.v3+json" }
Invoke-RestMethod -Uri "https://api.github.com/repos/iukimblock-collab/intax-dashboard/contents/index.html" -Headers $headers
```
