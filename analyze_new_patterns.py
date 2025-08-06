import pandas as pd
import os
import sys
import io
import re

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_hopyo_patterns(file_path, sheet_name):
    """호표 패턴과 구조를 더 자세히 분석"""
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
        
        # 호표 패턴 찾기
        hopyo_patterns = []
        
        # test1, est 스타일: "1.아스팔트포장깨기", "2.콘크리트깨기" 등
        pattern1 = r'^(\d+)\.\s*(.+)'
        
        # sgs 스타일: "산근 1 호표：", "산근 2 호표："
        pattern2 = r'산근\s*(\d+)\s*호표\s*[：:]'
        
        # ebs 스타일: "#1 토공", "#2 석축헐기"
        pattern3 = r'^#(\d+)\s+(.+)'
        
        print("\n[호표 패턴 검색]")
        for i in range(min(100, len(df))):  # 처음 100행만 검색
            for j in range(min(5, len(df.columns))):  # 처음 5열만 검색
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # 패턴1 검색
                    match1 = re.match(pattern1, cell_str)
                    if match1:
                        hopyo_num = match1.group(1)
                        work_name = match1.group(2).strip()
                        hopyo_patterns.append({
                            'row': i, 'col': j, 'type': 'dot_style',
                            'hopyo': hopyo_num, 'work': work_name,
                            'full': cell_str
                        })
                        print(f"[점 스타일] 행{i}: {hopyo_num}호표 - {work_name}")
                    
                    # 패턴2 검색
                    match2 = re.search(pattern2, cell_str)
                    if match2:
                        hopyo_num = match2.group(1)
                        # 작업명은 콜론 뒤에 있음
                        work_name = cell_str.split('：')[-1].strip() if '：' in cell_str else cell_str.split(':')[-1].strip()
                        hopyo_patterns.append({
                            'row': i, 'col': j, 'type': 'sanggun_style',
                            'hopyo': hopyo_num, 'work': work_name,
                            'full': cell_str
                        })
                        print(f"[산근 스타일] 행{i}: {hopyo_num}호표 - {work_name}")
                    
                    # 패턴3 검색
                    match3 = re.match(pattern3, cell_str)
                    if match3:
                        hopyo_num = match3.group(1)
                        work_name = match3.group(2).strip()
                        hopyo_patterns.append({
                            'row': i, 'col': j, 'type': 'hash_style',
                            'hopyo': hopyo_num, 'work': work_name,
                            'full': cell_str
                        })
                        print(f"[해시 스타일] 행{i}: {hopyo_num}호표 - {work_name}")
        
        # 헤더 위치와 구조 분석
        print("\n[헤더 구조 분석]")
        header_row = None
        for i in range(min(10, len(df))):
            row_text = ' '.join([str(df.iloc[i, j]) for j in range(len(df.columns)) if pd.notna(df.iloc[i, j])])
            if '산출근거' in row_text or '산 출 근 거' in row_text:
                header_row = i
                print(f"헤더 행 발견: {i}")
                # 헤더 행의 내용 출력
                headers = []
                for j in range(len(df.columns)):
                    if pd.notna(df.iloc[i, j]):
                        headers.append(f"{j}열: {df.iloc[i, j]}")
                print(f"헤더 내용: {headers[:10]}")  # 처음 10개만
                break
        
        # 데이터 구조 샘플링
        if hopyo_patterns:
            print("\n[첫 번째 호표의 데이터 구조]")
            first_hopyo = hopyo_patterns[0]
            start_row = first_hopyo['row']
            
            # 다음 호표까지의 범위 찾기
            end_row = start_row + 50  # 기본값
            if len(hopyo_patterns) > 1:
                end_row = hopyo_patterns[1]['row']
            
            print(f"\n{first_hopyo['hopyo']}호표 데이터 (행 {start_row} ~ {end_row}):")
            for i in range(start_row, min(start_row + 20, end_row, len(df))):
                row_data = []
                for j in range(min(8, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        row_data.append(f"{j}: {str(cell).strip()}")
                if row_data:
                    print(f"행{i}: {' | '.join(row_data)}")
        
        return hopyo_patterns
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        return []

def main():
    # 분석할 파일과 시트 목록
    files_to_analyze = [
        ('test1.xlsx', '단가산출_산근'),
        ('est.xlsx', '단가산출'),
        ('sgs.xls', '단가산출'),
        ('ebs.xls', '일위대가_산근')
    ]
    
    print("단가산출 시트 호표 패턴 분석")
    print("="*60)
    
    all_patterns = {}
    for file_name, sheet_name in files_to_analyze:
        if os.path.exists(file_name):
            patterns = analyze_hopyo_patterns(file_name, sheet_name)
            all_patterns[file_name] = patterns
        else:
            print(f"\n파일을 찾을 수 없음: {file_name}")
    
    # 요약
    print("\n\n" + "="*60)
    print("분석 요약")
    print("="*60)
    for file_name, patterns in all_patterns.items():
        if patterns:
            style = patterns[0]['type']
            count = len(patterns)
            print(f"\n{file_name}:")
            print(f"  - 스타일: {style}")
            print(f"  - 호표 개수: {count}")
            print(f"  - 첫 호표: {patterns[0]['hopyo']}호표 - {patterns[0]['work']}")

if __name__ == "__main__":
    main()