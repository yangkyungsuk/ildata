import pandas as pd
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def test_detection():
    """파일별 구조 상세 테스트"""
    
    files = [
        ('test1.xlsx', '일위대가_산근'),
        ('sgs.xls', '일위대가'),
        ('건축구조내역2.xlsx', '일위대가'),
        ('est.xlsx', '일위대가')
    ]
    
    for file_path, sheet_name in files:
        print(f"\n{'='*60}")
        print(f"{file_path} - {sheet_name}")
        print(f"{'='*60}")
        
        # 엑셀 읽기
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # 헤더 영역 확인
        print("\n헤더 영역 (0-5행):")
        for i in range(min(6, len(df))):
            row_info = []
            for j in range(min(6, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    row_info.append(f"{j}:{str(cell)[:20]}")
            if row_info:
                print(f"행{i}: {row_info}")
        
        # 호표 패턴 확인
        print("\n호표 패턴:")
        found_patterns = []
        
        for i in range(min(30, len(df))):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell) and '호표' in str(cell):
                    cell_str = str(cell)
                    
                    # 패턴 분류
                    if '：' in cell_str or ':' in cell_str:
                        pattern = "콜론형"
                    elif '(' in cell_str and ')' in cell_str:
                        pattern = "괄호형"
                    else:
                        pattern = "분리형"
                    
                    found_patterns.append({
                        'row': i,
                        'col': j,
                        'pattern': pattern,
                        'text': cell_str[:50]
                    })
                    
                    if len(found_patterns) >= 3:
                        break
            if len(found_patterns) >= 3:
                break
        
        for p in found_patterns:
            print(f"  행{p['row']} 열{p['col']}: {p['pattern']} - {p['text']}")
        
        # 헤더 키워드 확인
        print("\n헤더 키워드 위치:")
        keywords = ['호표', '품명', '공종', '규격', '단위', '수량']
        
        for keyword in keywords:
            for i in range(min(5, len(df))):
                for j in range(min(10, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell) and keyword in str(cell):
                        print(f"  '{keyword}' - 행{i} 열{j}")
                        break

if __name__ == "__main__":
    test_detection()