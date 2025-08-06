import pandas as pd
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

files = ['건축구조내역2.xlsx', '건축구조내역.xlsx']

for file in files:
    try:
        xls = pd.ExcelFile(file)
        print(f"\n{file} 시트 목록:")
        for i, sheet in enumerate(xls.sheet_names):
            print(f"  {i}: {sheet}")
            
        # 단가산출 관련 시트 찾기
        for sheet in xls.sheet_names:
            if '단가' in sheet or '산출' in sheet or '중기' in sheet:
                print(f"\n'{sheet}' 시트 샘플:")
                df = pd.read_excel(file, sheet_name=sheet, nrows=10, header=None)
                print(df)
                break
                
    except Exception as e:
        print(f"{file} 오류: {e}")