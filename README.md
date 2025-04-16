# Redmine SQA 우선순위 리포트 자동 생성기

Redmine 이슈를 기반으로 SQA 테스트 우선순위를 자동 계산하고, 우선순위에 따라 정렬된 Excel 보고서를 생성하는 자동화 스크립트입니다.


## 주요 기능

- Redmine API 연동을 통한 이슈 자동 수집
- 등록일과 우선순위 기준으로 테스트 우선도 판단 (🔺/⚠/⏳)
- Excel 리포트 자동 생성
  - 테스트 우선도별 색상 강조
  - 중앙 정렬, 열 너비 자동 조정
- `.bat` 파일 실행으로 더블클릭만으로 자동 리포트 생성 가능


## 실행 방법

1. Python 3 설치
2. 필수 라이브러리 설치:
```
pip install requests openpyxl matplotlib pandas
```
3. `Redmine_priority.py` 파일에서 `API_KEY` 를 본인의 Redmine 키로 변경
4. 스크립트 실행:
```
python Redmine_priority.py
```
또는 `run_priority_report.bat` 배치파일을 더블클릭하면 자동 실행


## 실행 결과

파일 이름은 자동으로 `SQA_Priority_Report_*.xlsx` 형식으로 저장  (하나의 워크북에 여러 시트가 포함)

- 프로젝트 이름, 전체 이슈 수, 잔여 이슈 수 요약
- 우선순위 기반 테스트 추천 목록 테이블
- 상태 ≠ Closed/Resolved 인 이슈만 포함
- 등록일과 우선순위 기준 자동 평가
- 테스트 우선도 (🔺 매우 높음 / ⚠ 보통 이상 / ⏳ 낮음)
- 우선도별 배경 색상 강조



## 파일 구조 예시

```
📁 Redmine_priority_report/
 ┣ 📄 Redmine_priority.py           → Redmine 이슈 기반 리포트 생성 스크립트
 ┣ 📄 run_priority_report.bat       → 파이썬 스크립트 자동 실행용 배치파일
 ┣ 📄 README.md                     → 프로젝트 설명 문서
 ┣ 📄 .gitignore                    → 자동 생성 리포트 제외 설정
 ┗ 📄 SQA_Priority_Report_*.xlsx    → 실행 시 자동 생성되는 우선순위 리포트 파일

```


## 👩‍💻 작성자
- 김예지 (SQA Engineer_
- GitHub: [@tmsnvl0415](https://github.com/tmsnvl0415)
