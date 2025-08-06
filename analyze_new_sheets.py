import pandas as pd
import os
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_sheet_structure(file_path, sheet_name):
    """특정 시트의 구조를 분석"""
    print(f"\n{'='*60}")
    print(f"파일: {file_path}")
    print(f"시트: {sheet_name}")
    print('='*60)
    
    try:
        # 파일 확장자에 따라 엔진 선택
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        
        print(f"시트 크기: {df.shape[0]}행 x {df.shape[1]}열")
        
        # 상위 20행 출력
        print("\n[상위 20행 데이터]")
        for i in range(min(20, len(df))):
            row_data = []
            for j in range(min(10, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    if cell_str:
                        row_data.append(f"{j}열: {cell_str}")
            if row_data:
                print(f"행 {i}: {' | '.join(row_data)}")
        
        # '산출근거' 헤더 찾기
        print("\n[산출근거 헤더 위치 찾기]")
        found_header = False
        for i in range(min(10, len(df))):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if pd.notna(cell) and '산출근거' in str(cell):
                    print(f"'산출근거' 발견: 행 {i}, 열 {j}")
                    found_header = True
                    
                    # 해당 행의 전체 헤더 출력
                    headers = []
                    for k in range(len(df.columns)):
                        header = df.iloc[i, k]
                        if pd.notna(header):
                            headers.append(f"{k}열: {str(header).strip()}")
                    print(f"헤더 행: {' | '.join(headers)}")
                    break
            if found_header:
                break
        
        # 숫자 패턴 찾기 (대분류/중분류)
        print("\n[숫자 패턴 분석]")
        number_pattern = []
        for i in range(len(df)):
            for j in range(min(3, len(df.columns))):  # 처음 3열만 확인
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    # 숫자만 있는 경우 확인
                    if cell_str.replace('.', '').isdigit():
                        number_pattern.append({
                            'row': i,
                            'col': j,
                            'value': cell_str,
                            'next_col_value': str(df.iloc[i, j+1]) if j+1 < len(df.columns) else ''
                        })
                        if len(number_pattern) <= 20:  # 처음 20개만 출력
                            print(f"행 {i}, 열 {j}: {cell_str} | 다음 열: {number_pattern[-1]['next_col_value'][:50]}")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        return False

def main():
    # 분석할 파일과 시트 목록
    files_to_analyze = [
        ('test1.xlsx', '단가산출_산근'),
        ('est.xlsx', '단가산출'),
        ('건축구조내역.xlsx', '중기단가산출'),  # 새 파일
        ('stmate.xlsx', '일위대가_산근'),  # 새 파일
        ('sgs.xls', '단가산출'),
        ('ebs.xls', '일위대가_산근')  # 새 파일
    ]
    
    print("단가산출 관련 시트 구조 분석")
    print("="*60)
    
    for file_name, sheet_name in files_to_analyze:
        if os.path.exists(file_name):
            analyze_sheet_structure(file_name, sheet_name)
        else:
            print(f"\n파일을 찾을 수 없음: {file_name}")

if __name__ == "__main__":
    main()