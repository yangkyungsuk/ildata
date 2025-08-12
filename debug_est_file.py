"""
est.xlsx 파일 상세 디버깅 및 분석
"""
import pandas as pd
import re

def debug_est():
    """est.xlsx 파일 상세 분석"""
    print("=" * 80)
    print("est.xlsx 파일 상세 분석")
    print("=" * 80)
    
    # 파일 읽기
    df = pd.read_excel('est.xlsx', sheet_name='일위대가', header=None)
    print(f"시트 크기: {df.shape}")
    
    # 1. 전체 데이터 구조 확인
    print("\n[처음 20행 데이터]")
    for i in range(min(20, len(df))):
        row_data = []
        for j in range(min(10, len(df.columns))):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                value = str(cell).strip()[:20]
                row_data.append(f"{j}:{value}")
        if row_data:
            print(f"행 {i:3}: {' | '.join(row_data)}")
    
    # 2. 호표 패턴 찾기
    print("\n[호표 패턴 찾기]")
    hopyo_patterns = [
        (r'제\s*(\d+)\s*호표', '제N호표'),
        (r'^(\d+)\.\s*(.+)', 'N. 작업명'),
        (r'^(\d+)\s+(.+)', 'N 작업명')
    ]
    
    found_hopyos = []
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                for pattern, pattern_name in hopyo_patterns:
                    match = re.match(pattern, cell_str)
                    if match:
                        # 작업명 찾기
                        work_name = ''
                        if pattern_name == '제N호표':
                            # 다음 셀 또는 다음 행에서 작업명 찾기
                            for k in range(j+1, min(j+5, len(df.columns))):
                                next_cell = df.iloc[i, k]
                                if pd.notna(next_cell):
                                    next_str = str(next_cell).strip()
                                    if next_str and not re.match(r'^[\d.,]+$', next_str):
                                        work_name = next_str
                                        break
                            
                            # 다음 행에서도 찾기
                            if not work_name and i+1 < len(df):
                                for k in range(len(df.columns)):
                                    next_cell = df.iloc[i+1, k]
                                    if pd.notna(next_cell):
                                        next_str = str(next_cell).strip()
                                        if next_str and not re.match(r'^[\d.,]+$', next_str) and '호표' not in next_str:
                                            work_name = next_str
                                            break
                        else:
                            # N. 패턴인 경우 그룹 2가 작업명
                            if len(match.groups()) > 1:
                                work_name = match.group(2)
                            else:
                                work_name = cell_str
                        
                        found_hopyos.append({
                            'row': i,
                            'col': j,
                            'pattern': pattern_name,
                            'num': match.group(1),
                            'work': work_name,
                            'full_text': cell_str[:50]
                        })
                        break
    
    print(f"발견된 호표: {len(found_hopyos)}개")
    for h in found_hopyos[:10]:
        print(f"  행 {h['row']:3}, 열 {h['col']}: [{h['pattern']}] 제{h['num']}호표 - {h['work'][:30]}")
    
    # 3. 실제 데이터 구조 분석
    if found_hopyos:
        print("\n[첫 번째 호표 상세 분석]")
        first_hopyo = found_hopyos[0]
        start_row = first_hopyo['row']
        end_row = found_hopyos[1]['row'] if len(found_hopyos) > 1 else min(start_row + 20, len(df))
        
        print(f"호표 범위: 행 {start_row} ~ {end_row}")
        print("\n산출근거 데이터:")
        
        for i in range(start_row, end_row):
            row_data = []
            for j in range(min(10, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    value = str(cell).strip()
                    row_data.append(value[:20])
            if row_data:
                print(f"  행 {i:3}: {' | '.join(row_data)}")
    
    # 4. 컬럼 구조 분석
    print("\n[컬럼 구조 분석]")
    # 헤더 찾기
    for i in range(min(10, len(df))):
        row_text = []
        for j in range(min(10, len(df.columns))):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                row_text.append(str(cell).strip())
        
        row_str = ' '.join(row_text).lower()
        if '품' in row_str or '규' in row_str or '단위' in row_str:
            print(f"헤더 후보 (행 {i}): {' | '.join(row_text[:8])}")
    
    return found_hopyos, df

if __name__ == "__main__":
    hopyos, df = debug_est()