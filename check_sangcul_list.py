import pandas as pd
import os
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def check_sangcul_list_sheet(file_path):
    """단가산출목록 관련 시트 확인"""
    try:
        xls = pd.ExcelFile(file_path)
        print(f"\n{'='*80}")
        print(f"파일: {file_path}")
        print(f"시트 목록: {xls.sheet_names}")
        
        # 단가산출목록 관련 시트 찾기
        list_sheets = []
        for sheet in xls.sheet_names:
            if '단가산출' in sheet and ('목록' in sheet or '대비' not in sheet):
                list_sheets.append(sheet)
        
        print(f"단가산출 관련 시트: {list_sheets}")
        
        # 각 목록 시트 확인
        for sheet_name in list_sheets:
            if '목록' in sheet_name:
                print(f"\n[{sheet_name}] 내용:")
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                print(f"크기: {df.shape}")
                
                # 상위 20행 출력
                print("\n상위 20행:")
                for i in range(min(20, len(df))):
                    row_data = []
                    for j in range(min(10, len(df.columns))):
                        cell = df.iloc[i, j]
                        if pd.notna(cell):
                            cell_str = str(cell).strip()
                            if cell_str:
                                row_data.append(f"{j}: {cell_str[:30]}")
                    if row_data:
                        print(f"행{i}: {' | '.join(row_data[:5])}")
                
                # 호표 관련 정보 찾기
                print("\n호표 정보 찾기:")
                hopyo_count = 0
                for i in range(len(df)):
                    for j in range(len(df.columns)):
                        cell = df.iloc[i, j]
                        if pd.notna(cell):
                            cell_str = str(cell).strip()
                            # 다양한 호표 패턴
                            if ('호표' in cell_str or '산근' in cell_str or 
                                (cell_str.isdigit() and 1 <= int(cell_str) <= 200)):
                                
                                # 같은 행의 다른 정보 출력
                                row_info = []
                                for k in range(len(df.columns)):
                                    c = df.iloc[i, k]
                                    if pd.notna(c):
                                        row_info.append(f"{k}: {str(c).strip()[:40]}")
                                
                                if len(row_info) > 2:  # 의미있는 데이터가 있는 경우
                                    print(f"행{i}: {' | '.join(row_info[:6])}")
                                    hopyo_count += 1
                                    if hopyo_count > 15:  # 처음 15개만
                                        print("...")
                                        break
                    if hopyo_count > 15:
                        break
                        
    except Exception as e:
        print(f"오류: {str(e)}")

# 분석할 파일들
files_to_check = [
    'test1.xlsx',
    'est.xlsx', 
    'sgs.xls',
    'ebs.xls',
    '건축구조내역.xlsx',
    '건축구조내역2.xlsx'
]

print("단가산출목록 시트 분석")
print("="*80)

for file_name in files_to_check:
    if os.path.exists(file_name):
        check_sangcul_list_sheet(file_name)
    else:
        print(f"\n파일 없음: {file_name}")