# 측정기 교정관리 (Calibration-manager)

교정대상 엑셀을 가져와 **입고 → 교정 → 출하** 전 과정을 관리하는 웹앱입니다.

- 화면: `index.html` (GitHub Pages로 실행)
- 서버: `gas/Code.js` (Google Apps Script + 스프레드시트) — main에 올라오면 자동 배포
- 버전: 웹앱 V1.0.0 / GAS V1.0

## 처음 설정 (한 번만)

1. **스프레드시트 만들기**: 구글 드라이브에서 새 스프레드시트(예: `교정관리DB`) 생성
2. **Apps Script 붙이기**: 확장 프로그램 → Apps Script → `gas/Code.js` 내용 붙여넣기 → 저장
3. **웹앱 배포**: 배포 → 새 배포 → 유형 *웹 앱* → 실행: *나*, 액세스: *모든 사용자* → 배포 → URL(…/exec) 복사
4. **자동 배포 Secrets** (저장소 Settings → Secrets and variables → Actions)
   - `CLASPRC_JSON`: scheduler 저장소와 같은 값
   - `GAS_SCRIPT_ID`: Apps Script → 프로젝트 설정 → 스크립트 ID
   - `GAS_DEPLOYMENT_ID`: 3번에서 만든 배포 ID(또는 웹앱 URL 전체)
5. **GitHub Pages**: Settings → Pages → Branch `main` / root
6. 웹앱 접속 → GAS URL 입력 → **가입(첫 가입자 = 관리자)** → [관리] 탭에서 교정대상 엑셀 가져오기
7. (선택) 접속 비밀번호: Apps Script → 프로젝트 설정 → 스크립트 속성 `ACCESS_CODE`

## 권한

| 권한 | 할 수 있는 일 |
|---|---|
| 관리자 | 엑셀 가져오기, 사용자 관리, 교정방식 변경, 단계 되돌리기, 삭제 |
| 담당자 | 입고·교정·출하 처리(일괄 포함), 비고 |
| 뷰어 | 조회 |

## 스프레드시트 탭 (자동 생성)

`교정대상` · `교정이력` · `교정완료기록`(이전 회차 보관) · `사용자` · `설정`

※ 측정기 데이터는 스프레드시트에만 저장되고 이 저장소에는 올라가지 않습니다.
