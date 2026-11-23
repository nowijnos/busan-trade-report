부산 수출입동향 자동작성 웹사이트 v32 — 업로드 안내

1. 이 ZIP의 압축을 풀어주세요.
2. GitHub 저장소 nowjinos/busan-trade-report → Add file → Upload files.
3. 압축 해제한 내용물(engine.py, server.py, kstat_collect.py, xlreader.py, requirements.txt, static, templates)을 저장소 루트에 같은 폴더 구조로 올려주세요.
4. Commit changes → Render 배포 성공을 확인해 주세요.
5. 웹사이트에서 Ctrl+Shift+R 새로고침.
6. 웹사이트 '사용방법'에서 도우미 ZIP을 내려받아 압축 해제 후 chrome://extensions에서 설치하세요.

이번 수정의 핵심은 엑셀 내부에 '지역 : 전국'으로 표시되더라도 본부 부산 수출/수입 총계(또는 중국 수입 총계)와 정확히 일치할 때만 자료를 인정하는 것입니다. 잘못된 타지역 파일은 제외합니다.
기존 HWPX 템플릿과 xlreader.py, 그래프 생성 로직은 그대로 유지했습니다.
