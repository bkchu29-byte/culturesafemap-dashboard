# 컬처세이프맵 — Colab 없이 실행

## 사용

1. 저장소 **Actions → CultureSafeMap HTML·한글 공문 갱신 → Run workflow → main → Run workflow**.
2. 실행이 성공하면 루트 `index.html`만 갱신됩니다. 기존 GitHub → Cloudflare 연동을 그대로 사용합니다.
3. 대시보드의 공문 화면에서 공연을 선택하고 초안을 연 뒤 **한글(HWPX) 내려받기**를 누릅니다.
4. 실행 결과 하단 `CultureSafeMap-HTML-HWPX`에서도 생성 HTML, 검증 보고서, 한글 공문 예시 한 건을 받을 수 있습니다.

Colab과 개인 GitHub 토큰은 필요하지 않습니다. Actions의 기본 `GITHUB_TOKEN`을 사용합니다.
최초 설치 및 자동화 코드/템플릿 변경 시에도 실행됩니다. 예약 실행은 아직 설정하지 않았습니다.

## 변경 범위

- 공문 판단: 기존 자동화 v4의 과거 패턴·최근 동향·업무 우선순위를 그대로 사용합니다.
- 입력: `Yu-Hyemin/culture_safe_map` main의 output CSV 3개를 읽기만 합니다. 실행 중에는 같은 커밋의 파일을 사용합니다.
- 웹페이지: 접속할 때 main의 `disease.csv`, `ili_region.csv`, `performance_master.csv`를 다시 불러오는 기존 연결을 유지합니다.
- 공문 판단 결과는 Actions 실행 때 갱신됩니다. 원본 CSV가 바뀌었다면 공문 사용 전에 다시 실행하세요.
- 화면: 저장소의 승인 v29를 기준으로 합니다. 지도·색상·범례·필터·팝업·세 문서 미리보기 CSS는 변경하지 않았습니다.
- 한글: v30의 별도 고정 HWPX 템플릿에 선택 공연의 데이터를 넣습니다. HTML CSS 또는 Word 디자인 코드를 공유하지 않습니다.

## 고정 파일

- `automation/templates/dashboard.html`: 승인 화면과 HWPX 고정 서식. 자동화는 이 파일의 CSS를 수정하지 않습니다.
- `automation/notice_core_v4.py`: 원래 v4 판단 코드. 일괄 Word 예시는 비활성화(`MAX_DOCS=0`)하고 한국 날짜를 사용합니다.
- `automation/build_dashboard.py`: 데이터 다운로드, 판단 실행, 공문 데이터 바인딩, HWPX 예시 구조 검사.
- `automation/test_dashboard.cjs`: 접속 시 로딩·공연 선택·문서 데이터·기존 CSS·브라우저 HWPX 생성 검사.

## 유의사항

HWPX는 한글에서 편집 가능한 문서 형식이며, 구형 바이너리 `.hwp` 파일은 아닙니다.
HWPX ZIP/XML 구조와 데이터 바인딩을 자동 검사합니다. 실제 한글 프로그램의 글꼴·쪽 나눔은 발송 전에 최종 확인해야 합니다.
기관 전자결재·이메일 발송은 실행하지 않으며 초안 생성까지만 연결합니다.

`index.html`은 생성 결과입니다. 승인된 화면을 수정하려면 고정 템플릿을 해당 요청 범위만 수정하세요.
생성 도중 main이 변경되면 강제 덮어쓰지 않고 push가 실패합니다. 최신 main에서 다시 실행하세요.
