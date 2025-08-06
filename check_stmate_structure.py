"""
stmate.xlsx 파일 구조 분석
"""
import pandas as pd
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_stmate_file():
    """stmate.xlsx 파일 구조 분석"""
    file_path = 'stmate.xlsx'
    
    try:
        # 시트 목록 먼저 확인
        try:
            xls = pd.ExcelFile(file_path)
            sheets = xls.sheet_names
            print(f"시트 목록: {sheets}")
        except Exception as e:
            print(f"시트 목록 읽기 오류: {e}")
            # openpyxl로 직접 시도
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheets = wb.sheetnames
            print(f"시트 목록 (openpyxl): {sheets}")
        
        # 각 시트 구조 분석
        for sheet_name in sheets:
            print(f"\n{'='*50}")
            print(f"시트: {sheet_name}")
            print('='*50)
            
            try:
                # 처음 20행 읽기
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
                print(f"크기: {df.shape}")
                
                # 내용 확인
                print("\n처음 10행 내용:")
                for i in range(min(10, len(df))):
                    row_content = []
                    for j in range(min(10, len(df.columns))):
                        cell = df.iloc[i, j]
                        if pd.notna(cell):
                            content = str(cell).strip()
                            if content:
                                row_content.append(f"[{j}]:{content[:30]}")
                    if row_content:
                        print(f"행{i}: {' | '.join(row_content)}")
                
                # 호표 패턴 찾기
                print(f"\n호표 패턴 검색:")
                patterns = [
                    r'제\d+호표',      # 제N호표
                    r'호표\s*\d+',     # 호표 N
                    r'\d+\.\s*',       # N. 형식
                    r'산근\s*\d+',     # 산근 N
                    r'#\d+',           # #N
                    r'\(\s*산근\s*\d+\s*\)'  # (산근 N)
                ]
                
                import re
                found_patterns = []
                
                for i in range(len(df)):
                    for j in range(len(df.columns)):
                        cell = df.iloc[i, j]
                        if pd.notna(cell):
                            cell_str = str(cell).strip()
                            for pattern in patterns:
                                if re.search(pattern, cell_str):
                                    found_patterns.append({
                                        'row': i,
                                        'col': j,
                                        'pattern': pattern,
                                        'text': cell_str[:50]
                                    })
                
                if found_patterns:
                    print("발견된 호표 패턴:")
                    for p in found_patterns[:5]:  # 처음 5개만
                        print(f"  행{p['row']}, 열{p['col']}: {p['pattern']} -> {p['text']}")
                else:
                    print("호표 패턴을 찾을 수 없습니다.")
                
            except Exception as e:
                print(f"시트 '{sheet_name}' 분석 오류: {e}")
    
    except Exception as e:
        print(f"파일 분석 오류: {e}")

if __name__ == "__main__":
    analyze_stmate_file()