# doc-to-markdown

HWP·HWPX·DOCX·PDF·PPTX·XLSX 같은 문서를 Claude가 마크다운으로 바꾸게 하는 Claude 스킬입니다.

- 원본은 읽기만 하고 건드리지 않습니다.
- 같은 문서(같은 SHA-256)는 다시 바꾸지 않습니다.
- 변환 기록을 해시 사슬로 남기고, 확인하지 못한 점(서식·그림·표 모양 등)을 결과에 함께 적습니다.
- 개인정보·AI 지시문 모양은 개수만 알립니다.

## 받기와 설치

1. [doc-to-markdown-0.2.0-public.zip 받기](https://github.com/10baggerbrain-code/doc-to-markdown/raw/main/dist/doc-to-markdown-0.2.0-public.zip) (`dist/` 폴더). 압축은 풀지 않습니다.
   SHA-256: `7ce3437087f103d206f14be71c90a99c5dd3f38ba3449e4a4e0d99564bc310f3`
2. Claude에서 **사용자 정의(Customize) → 스킬(Skills) → + → 스킬 만들기 → 스킬 업로드**로 ZIP을 올립니다.
3. 켠 뒤 "이 문서들 마크다운으로 바꿔 줘"처럼 말하면 됩니다.

스킬은 코드 실행이 켜져 있어야 동작합니다. 메뉴 이름은 Claude 도움말 센터 기준이며 바뀔 수 있습니다.

## 고쳐 쓰기

- `doc-to-markdown/SKILL.md`: Claude가 읽는 지침(언제 쓰는지, 보고 방식)
- `doc-to-markdown/scripts/doc2md.py`: 변환과 안전 규칙
- 고친 뒤에는 `python3 -I doc-to-markdown/tests/run_tests.py`가 모두 통과해야 합니다(현재 21개).

## 알림

- 글자 위주 변환입니다. 서식·그림·표 모양, 스캔 PDF 내용, 엑셀 수식 자체는 옮기지 않습니다.
- 시험은 직접 만든 견본으로 돌린 것입니다. 실제 한글 프로그램으로 만든 HWP 파일로는 아직 충분히 검증하지 않았습니다.
- 본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
- 이 저장소의 코드와 글은 AI(Claude)가 작성했습니다. 보증 없이 제공합니다.

## 라이선스

MIT — `LICENSE` 참고. Copyright (c) 2026 세린 유니버스 창조자
