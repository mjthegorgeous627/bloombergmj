workbench ver.2 (2026-10-14) 재료.
- patch_A.py + patch_layout.py + A_09_side_panel.js + A_extra.css : 현재 templates/ 를 복사해 v2 화면을 만든다.
  V2_DIR=<sap_automation>\templates_v2 python patch_A.py   (templates_v2/ 를 다시 만들어 줌 - 서버는 /v2 에서 요청마다 새로 조립)
- build_preview.py : DB 사본으로 미리보기 단일 HTML 생성(서버 호출은 가짜 응답)
- pwH/pwI/pwJ.py : Playwright 검증(원본 vs v2 동작 비교, 버튼 호출 확인)
v1 은 원래 주소(/), v2 는 /v2 . 같은 서버·같은 DB. 서버 추가분: v2_ops.py (빠른메모/오더번호 수정/OBD·티켓 바로가기).
