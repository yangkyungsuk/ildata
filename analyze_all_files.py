"""
모든 엑셀 파일의 일위대가 시트 구조 분석
각 파일의 패턴과 컬럼 구조를 파악
"""
import pandas as pd
import re
import os

def analyze_file_structure(file_path, sheet_name):
    """파일의 일위대가 시트 구조 분석"""
    print(f"\n{'='*70}")
    print(f"파일: {file_path}")
    print(f"시트: {sheet_name}")
    print('='*70)
    
    try:
        # 파일 읽기
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        print(f"시트 크기: {df.shape[0]}행 x {df.shape[1]}열")
        
        # 1. 호표 패턴 찾기
        print("\n[호표 패턴 분석]")
        hopyo_patterns = [
            (r'제\s*(\d+)\s*호표', '제N호표'),
            (r'^#(\d+)\s+(.+)', '#N 작업명'),
            (r'^No\.(\d+)\s+(.+)', 'No.N 작업명'),
            (r'(\d+)\.\s*(.+)', 'N. 작업명'),
            (r'산근\s*(\d+)\s*호표', '산근N호표'),
            (r'\(\s*산근\s*(\d+)\s*\)', '(산근N)')
        ]
        
        found_patterns = []
        hopyo_list = []
        
        for i in range(min(100, len(df))):  # 처음 100행만 검사
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    for pattern, pattern_name in hopyo_patterns:
                        match = re.match(pattern, cell_str)
                        if match:
                            if pattern_name not in found_patterns:
                                found_patterns.append(pattern_name)
                            hopyo_list.append({
                                'row': i,
                                'col': j,
                                'pattern': pattern_name,
                                'text': cell_str[:50]
                            })
                            break
        
        print(f"발견된 패턴: {found_patterns}")
        print(f"호표 개수: {len(hopyo_list)}개")
        
        if hopyo_list:
            print("\n처음 5개 호표:")
            for h in hopyo_list[:5]:
                print(f"  행 {h['row']+1}, 열 {h['col']+1}: {h['text']}")
        
        # 2. 헤더/컬럼 구조 분석
        print("\n[컬럼 구조 분석]")
        
        # 품명, 규격, 단위, 수량 관련 키워드 찾기
        keywords = {
            '품명': ['품명', '품 명', '자재명', '명칭', '공종', '품목'],
            '규격': ['규격', '규 격', '사양', 'SPEC'],
            '단위': ['단위', '단 위', 'UNIT'],
            '수량': ['수량', '수 량', '물량', 'QTY']
        }
        
        header_rows = []
        for i in range(min(20, len(df))):
            row_text = []
            for j in range(min(10, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    row_text.append(str(cell).strip())
            
            # 키워드가 포함된 행 찾기
            for category, words in keywords.items():
                for word in words:
                    if any(word in text for text in row_text):
                        header_rows.append({
                            'row': i,
                            'category': category,
                            'text': ' | '.join(row_text[:5])
                        })
                        break
        
        if header_rows:
            print("헤더 후보 행:")
            unique_rows = {}
            for h in header_rows:
                if h['row'] not in unique_rows:
                    unique_rows[h['row']] = h
                    print(f"  행 {h['row']+1}: {h['text']}")
        
        # 3. 데이터 샘플 분석
        print("\n[데이터 샘플]")
        if hopyo_list:
            # 첫 번째 호표 다음 5행 출력
            first_hopyo_row = hopyo_list[0]['row']
            print(f"첫 번째 호표(행 {first_hopyo_row+1}) 다음 데이터:")
            
            for i in range(first_hopyo_row + 1, min(first_hopyo_row + 6, len(df))):
                row_data = []
                for j in range(min(8, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        value = str(cell).strip()
                        if value:
                            row_data.append(value[:20])
                if row_data:
                    print(f"  행 {i+1}: {' | '.join(row_data)}")
        
        # 4. 컬럼 위치 추정
        print("\n[컬럼 위치 추정]")
        if hopyo_list and len(hopyo_list) > 0:
            # 첫 번째 호표 이후 데이터 분석
            start_row = hopyo_list[0]['row'] + 1
            end_row = min(start_row + 10, len(df))
            
            col_types = {}
            for col_idx in range(min(10, len(df.columns))):
                col_values = []
                for row_idx in range(start_row, end_row):
                    cell = df.iloc[row_idx, col_idx]
                    if pd.notna(cell):
                        col_values.append(str(cell).strip())
                
                if col_values:
                    # 컬럼 타입 추정
                    if any(re.match(r'^[가-힣]+', v) for v in col_values):
                        if col_idx not in col_types or '품명' not in col_types:
                            col_types[col_idx] = '품명 추정'
                    elif any(v in ['인', '%', 'M3', 'M2', 'M', 'KG', 'TON', '개', '대', 'HR', '조', 'm3'] for v in col_values):
                        col_types[col_idx] = '단위 추정'
                    elif any(re.match(r'^\d+\.?\d*$', v) for v in col_values):
                        # 작은 숫자는 수량, 큰 숫자는 금액
                        nums = [float(v) for v in col_values if re.match(r'^\d+\.?\d*$', v)]
                        if nums and max(nums) < 1000:
                            col_types[col_idx] = '수량 추정'
                        elif nums and max(nums) > 10000:
                            col_types[col_idx] = '금액 추정'
            
            for col_idx, col_type in sorted(col_types.items()):
                print(f"  컬럼 {col_idx}: {col_type}")
        
        return {
            'file': file_path,
            'sheet': sheet_name,
            'shape': df.shape,
            'patterns': found_patterns,
            'hopyo_count': len(hopyo_list),
            'hopyo_list': hopyo_list[:10]  # 처음 10개만
        }
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        return None

def main():
    """모든 파일 분석"""
    files_to_analyze = [
        ('est.xlsx', '일위대가'),
        ('sgs.xls', '일위대가'),
        ('건축구조내역.xlsx', '일위대가'),
        ('건축구조내역2.xlsx', '일위대가')
    ]
    
    print("모든 일위대가 시트 구조 분석")
    print("=" * 70)
    
    results = []
    for file_path, sheet_name in files_to_analyze:
        if os.path.exists(file_path):
            result = analyze_file_structure(file_path, sheet_name)
            if result:
                results.append(result)
        else:
            print(f"\n파일 없음: {file_path}")
    
    # 요약
    print("\n" + "=" * 70)
    print("분석 요약")
    print("=" * 70)
    
    for result in results:
        print(f"\n{result['file']}:")
        print(f"  - 시트: {result['sheet']}")
        print(f"  - 크기: {result['shape']}")
        print(f"  - 패턴: {result['patterns']}")
        print(f"  - 호표 수: {result['hopyo_count']}")

if __name__ == "__main__":
    main()