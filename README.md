# 한국 건설 견적서 Excel 파서 (Korean Construction Estimate Excel Parser)

한국 건설업계에서 사용하는 일위대가 견적서 Excel 파일을 자동으로 파싱하는 프로그램입니다.

## 🚀 주요 기능

- **자동 파일 구조 감지**: 업로드된 Excel 파일의 구조를 자동으로 분석
- **4가지 파일 형식 지원**:
  - TEST1 형식: 호표가 별도 열에 분리된 형식
  - SGS 형식: 콜론(:)으로 구분된 통합 형식
  - 건축구조 형식: 괄호 안에 호표가 포함된 형식
  - EST 형식: 호표가 1열에 위치한 형식
- **JSON 출력**: 구조화된 데이터를 JSON 형식으로 저장

## 📋 필요 사항

- Python 3.7 이상
- pandas
- openpyxl
- xlrd

## 🛠️ 설치 방법

1. 저장소 클론:
```bash
git clone https://github.com/[username]/korean-construction-excel-parser.git
cd korean-construction-excel-parser
```

2. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

## 📖 사용 방법

### 1. 스마트 파서 (자동 감지)

파일 구조를 자동으로 감지하고 적절한 파서를 선택합니다:

```bash
# 특정 파일 파싱
python smart_parser.py your_file.xlsx

# 테스트 파일 모두 파싱
python smart_parser.py
```

### 2. 개별 파서 사용

각 파일 형식에 맞는 전용 파서를 직접 사용할 수도 있습니다:

```bash
python parser_test1.py      # TEST1 형식
python parser_sgs.py        # SGS 형식
python parser_construction.py  # 건축구조 형식
python parser_est.py        # EST 형식
```

## 📁 프로젝트 구조

```
korean-construction-excel-parser/
│
├── smart_parser.py           # 메인 자동 파서
├── parser_test1.py          # TEST1 형식 전용 파서
├── parser_sgs.py            # SGS 형식 전용 파서
├── parser_construction.py    # 건축구조 형식 전용 파서
├── parser_est.py            # EST 형식 전용 파서
├── check_all_headers.py     # 파일 구조 분석 도구
├── verify_all_results.py    # 결과 검증 도구
│
├── test_files/              # 테스트 파일 (별도 제공)
│   ├── test1.xlsx
│   ├── sgs.xls
│   ├── 건축구조내역2.xlsx
│   └── est.xlsx
│
└── README.md
```

## 📊 출력 형식

파싱된 데이터는 다음과 같은 JSON 구조로 저장됩니다:

```json
{
  "file": "파일명.xlsx",
  "sheet": "시트명",
  "detected_type": "파일타입",
  "hopyo_count": 12,
  "hopyo_data": {
    "호표1": {
      "호표번호": 1,
      "작업명": "철근 현장가공",
      "규격": "Type-Ⅰ",
      "단위": "TON",
      "세부항목": [
        {
          "품명": "철근공",
          "규격": "일반공사 직종",
          "단위": "인",
          "수량": 0.67,
          "비고": ""
        }
      ],
      "항목수": 3
    }
  }
}
```

## 🔍 지원되는 Excel 구조

### 1. TEST1 형식
```
호표 | 품명 | 규격 | 단위 | 수량 | ... | 비고
제1호표 | 철근 현장가공 | Type-Ⅰ | TON | ...
```

### 2. SGS 형식
```
공종 | 규격 | 수량 | 단위 | ...
제1호표：탄성포장 철거 ㎡ 당
특별인부 | | 5 | 인 | ...
```

### 3. 건축구조 형식
```
품명 | 규격 | 단위 | 수량 | ...
M-101 정밀여과장치 설치 ... (호표 1)
일반기기설치 | 1~5TON 미만 | TON | 2.75 | ...
```

### 4. EST 형식
```
| 호표 | 품명 | 규격 | 수량 | 단위 |
| 제1호표 | 특수지역 위험목 제거 | 20cm이하 | | 본 |
| | 벌목부 | | 0.28 | 인 |
```

## ⚠️ 주의사항

- Excel 파일은 UTF-8 인코딩을 지원합니다
- 일위대가 관련 시트만 처리합니다 (목록표, 총괄표 제외)
- 파일명에 한글이 포함되어도 정상 작동합니다

## 🤝 기여 방법

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.

## 👥 개발자

- 개발: [Your Name]
- 문의: [Your Email]

## 🙏 감사의 말

한국 건설업계의 디지털화에 기여하고자 하는 모든 분들께 감사드립니다.