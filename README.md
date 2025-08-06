# 한국 건설업체 엑셀 단가산출서 파싱 시스템

## 개요
한국 건설업체에서 사용하는 다양한 형태의 엑셀 단가산출서를 JSON으로 변환하는 파싱 시스템입니다.
호표(工票) 사이의 모든 행을 순서대로 완전히 보존하여 데이터 누락 없이 변환합니다.

## 주요 기능
- **완전한 순서 보존**: 호표와 호표 사이의 모든 행을 원본 순서대로 보존
- **빈 행 포함**: 순서 유지를 위해 빈 행도 모두 저장
- **다양한 호표 패턴 지원**: 제N호표, #N, No.N, 산근N, (산근N) 등
- **목록 시트 검증**: 목록표와 실제 파싱 결과 검증
- **UTF-8 완전 지원**: 한국어 완벽 처리

## 지원 파일 형식

### 1. test1.xlsx - 단가산출_산근 시트
- **호표 패턴**: `N. 작업명` 형식
- **목록 시트**: 단가산출목록표 (`제N호표` 형식)
- **파서**: `parser_sangcul_test1_v4.py`
- **결과**: 12개 호표 파싱

### 2. sgs.xls - 단가산출 시트  
- **호표 패턴**: `산근 N 호표 : 작업명` 형식
- **목록 시트**: 단가산출 목록 (공종 헤더)
- **파서**: `parser_sangcul_sgs_v4.py`
- **결과**: 102개 호표 파싱

### 3. est.xlsx - 단가산출 시트
- **호표 패턴**: `N. 작업명` 형식 (col 1부터 시작)
- **목록 시트**: 단가산출총괄표
- **파서**: `parser_sangcul_est_v4.py` 
- **결과**: 11개 호표 파싱

### 4. ebs.xls - 일위대가_산근 시트
- **호표 패턴**: `#N 작업명` 형식
- **목록 시트**: 일위대가목록 (`#N` 형식)
- **파서**: `parser_sangcul_ebs_v4.py`
- **결과**: 765개 호표 파싱

### 5. 건축구조내역.xlsx - 중기단가산출서 시트
- **호표 패턴**: `작업명 (산근 N)` 형식
- **목록 시트**: 자동 탐지 (목록/총괄/중기단가 시트)
- **파서**: `parser_sangcul_construction_v4.py`

### 6. stmate.xlsx - 복합 구조 (2개 파서)
#### 6-1. 단가산출 구조 (일위대가_산근 시트)
- **호표 패턴**: `#N 작업명` 형식
- **목록 시트**: 일위대가목록 (`#N` 형식)
- **파서**: `parser_sangcul_stmate_v4.py`
- **결과**: 32개 호표 파싱
- **특이사항**: xlwings 라이브러리 사용 (openpyxl 호환성 문제)

#### 6-2. 일위대가 구조 (일위대가_호표 시트)
- **호표 패턴**: `No.N 작업명` 형식  
- **목록 시트**: 일위대가목록 (일위대가 항목들)
- **파서**: `parser_ilwidae_stmate_v4.py`
- **결과**: 46개 호표 파싱
- **특이사항**: xlwings 라이브러리 사용

## 사용 방법

### 1. 개별 파일 파싱
```bash
python parser_sangcul_test1_v4.py
python parser_sangcul_sgs_v4.py  
python parser_sangcul_est_v4.py
python parser_sangcul_ebs_v4.py
python parser_sangcul_construction_v4.py
python parser_sangcul_stmate_v4.py
python parser_ilwidae_stmate_v4.py
```

### 2. 스마트 파서 (자동 파일 인식)
```bash
# 특정 파일 파싱
python smart_sangcul_parser.py test1.xlsx

# 모든 파일 일괄 파싱  
python smart_sangcul_parser.py --all
```

### 3. 필요한 라이브러리
```bash
pip install pandas xlrd openpyxl xlwings
```

## JSON 출력 구조

### 기본 구조
```json
{
  "file": "파일명.xlsx",
  "sheet": "시트명", 
  "total_hopyo_count": 12,
  "validation": {
    "list_count": 12,
    "parsed_count": 12,
    "match": true
  },
  "hopyo_data": {
    "호표1": {
      "호표번호": "1",
      "작업명": "아스팔트포장깨기(대형브레이커+0.7㎥)[㎥]",
      "시작행": 2,
      "종료행": 50,
      "총_행수": 49,
      "내용있는_행수": 30,
      "모든_행_데이터": [...]
    }
  }
}
```

### 모든_행_데이터 구조 (v4 핵심 기능)
```json
"모든_행_데이터": [
  {
    "row_number": 2,
    "columns": {
      "col_0": "1.아스팔트포장깨기(대형브레이커+0.7㎥)[㎥]",
      "col_1": "3770.0",
      "col_2": "1033.0"
    },
    "has_content": true
  },
  {
    "row_number": 3,
    "columns": {},
    "has_content": false
  }
]
```

## 파일 구조
```
ildata/
├── README.md                              # 이 파일
├── smart_sangcul_parser.py                 # 스마트 자동 파서
├── 
├── # v4 파서들 (완전한 순서 보존)
├── parser_sangcul_test1_v4.py             # test1.xlsx 파서
├── parser_sangcul_sgs_v4.py               # sgs.xls 파서  
├── parser_sangcul_est_v4.py               # est.xlsx 파서
├── parser_sangcul_ebs_v4.py               # ebs.xls 파서
├── parser_sangcul_construction_v4.py      # 건축구조내역.xlsx 파서
├── parser_sangcul_stmate_v4.py            # stmate.xlsx 단가산출 파서
├── parser_ilwidae_stmate_v4.py            # stmate.xlsx 일위대가 파서
├──
├── # 원본 엑셀 파일들
├── test1.xlsx
├── sgs.xls
├── est.xlsx  
├── ebs.xls
├── 건축구조내역.xlsx
├── 건축구조내역2.xlsx
├── stmate.xlsx
├──
├── # 생성된 JSON 결과 파일들
├── test1_sangcul_parsed_v4.json
├── sgs_sangcul_parsed_v4.json
├── est_sangcul_parsed_v4.json
├── ebs_sangcul_parsed_v4.json
├── construction_sangcul_parsed_v4.json
├── stmate_sangcul_parsed_v4.json
└── stmate_ilwidae_parsed_v4.json
```

## v4 파서의 핵심 개선사항

### 1. 완전한 순서 보존  
- **문제**: 기존 v3 파서는 데이터를 분류하여 저장했으나 원본 행 순서가 손실됨
- **해결**: `모든_행_데이터` 배열에 호표 사이의 모든 행을 순서대로 저장
- **중요성**: "호표와 호표사이 정확히 행에 데이터에 순서에 맞게 json으로 변환"

### 2. 완전성 보장
- **문제**: "하나에 내용이라도 빠지면 심각한 오류를 범할수있음"
- **해결**: 빈 행까지 포함하여 모든 행을 저장 (`has_content: false`로 표시)
- **검증**: row_number로 원본 엑셀 행 번호 추적 가능

### 3. 통합 추출 함수
```python
def extract_all_rows_in_order(df, start_row, end_row):
    """호표 사이의 모든 행을 순서대로 완전히 추출"""
    all_rows = []
    
    for row_idx in range(start_row, end_row):
        row_data = {
            'row_number': row_idx,
            'columns': {},
            'has_content': False
        }
        
        # 모든 컬럼 검사 (최대 20컬럼)
        for col_idx in range(min(20, len(df.columns))):
            cell = df.iloc[row_idx, col_idx]
            if pd.notna(cell):
                content = str(cell).strip()
                if content:
                    row_data['columns'][f'col_{col_idx}'] = content
                    row_data['has_content'] = True
        
        all_rows.append(row_data)
    
    return all_rows
```

## 특별 고려사항

### 1. stmate.xlsx 파일
- **문제**: openpyxl/pandas로 읽을 수 없는 custom properties 오류
- **해결**: xlwings 라이브러리 사용
- **설치**: `pip install xlwings` (Excel이 설치된 환경 필요)

### 2. 인코딩 처리
- **UTF-8 설정**: 모든 파서에서 한국어 완벽 지원
```python
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
```

### 3. 호표 패턴별 정규식
- `제N호표`: `r'제(\\d+)호표'`
- `#N 작업명`: `r'^#(\\d+)\\s+(.+)'`  
- `No.N 작업명`: `r'^No\\.(\\d+)\\s+(.+)'`
- `N. 작업명`: `r'^(\\d+)\\.\\s*(.+)'`
- `산근 N 호표`: `r'산근\\s*(\\d+)\\s*호표\\s*[：:](.+)'`
- `(산근 N)`: `r'(.+)\\(\\s*산근\\s*(\\d+)\\s*\\)'`

## 버전 히스토리
- **v1**: 기본 파싱 기능
- **v2**: 목록 시트 검증 추가  
- **v3**: 세부 데이터 분류 수집
- **v4**: 완전한 순서 보존 (현재 버전)

## 개발자 노트
- 모든 파서는 동일한 구조와 출력 형식 사용
- `smart_sangcul_parser.py`가 파일명으로 자동 인식
- JSON 출력은 UTF-8로 저장되며 `ensure_ascii=False` 사용
- 각 파서는 독립적으로 실행 가능하도록 설계

## 문제 해결

### Q: 파싱 결과가 없거나 오류 발생
A: 1) 파일 경로 확인 2) 시트 이름 확인 3) 호표 패턴 확인

### Q: stmate 파일 파싱 오류  
A: xlwings 설치 확인: `pip install xlwings`

### Q: 한국어 깨짐
A: UTF-8 인코딩 설정 확인

---
*이 시스템은 한국 건설업계의 다양한 단가산출서 형식을 표준화된 JSON으로 변환하여 데이터 분석 및 검증 프로그램 개발을 지원합니다.*