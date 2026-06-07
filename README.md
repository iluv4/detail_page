# 겟머슬 11세대 RPM LED 자이로볼 — 상세페이지

쿠팡 / 스마트스토어용 모바일 상세페이지 (15장, 각 780×2000 PNG).

## 폴더 구조
- `assets/reference/` — 제품 사진·로고 등 생성 레퍼런스 원본
- `output/` — 생성된 상세페이지 PNG (1.png ~ 15.png)
- `scripts/` — 후처리(정확한 780×2000 리사이즈) 유틸
- `docs/` — 15장 카피/구성 기획안

## 진행 방식
1. 레퍼런스(제품 사진 + GetMuscle 로고) 확보
2. nano_banana_pro 모델로 장별 생성
3. 정확한 780×2000 규격으로 후처리

금지 표현/대체 표현 가이드는 `docs/copy-guide.md` 참고.

## 산출물 (Word 문서)
- `docs/겟머슬_상세페이지_기획안.docx` — 15장 구매전환 상세페이지 최종 기획안 (비교표·루틴표·강조 박스 포함)
- `docs/겟머슬_인플루언서_협업_브리프.docx` — 운동 인플루언서 협업 브리프

생성 스크립트(원본은 `docs/`의 텍스트 기획안):
```bash
pip install python-docx
python3 scripts/build_plan_docx.py   # 상세페이지 기획안 .docx
python3 scripts/build_docx.py        # 인플루언서 브리프 .docx
```
