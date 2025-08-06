import pandas as pd
import os
import sys
import io
import re

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_hopyo_data_range(file_path, sheet_name):
    """각 호표 사이의 데이터 범위를 분석"""
    print(f"\n{'='*80}")
    print(f"파일: {file_path} | 시트: {sheet_name}")
    print('='*80)
    
    try:
        # 파일 읽기
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        
        # 패턴 정의
        patterns = [
            (r'^(\d+)\.\s*(.+)', 'dot_style'),  # 1.작업명
            (r'산근\s*(\d+)\s*호표\s*[：:]', 'sanggun_style'),  # 산근 1 호표
            (r'^#(\d+)\s+(.+)', 'hash_style'),  # #1 작업명
            (r'^제\s*(\d+)\s*호표', 'je_style'),  # 제1호표 (기존 일위대가)
            (r'호표\s*(\d+)', 'hopyo_style')  # 호표 1 (기존 일위대가)
        ]
        
        # 호표 찾기
        hopyo_list = []
        for i in range(len(df)):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    for pattern, style in patterns:
                        match = re.search(pattern, cell_str)
                        if match:
                            hopyo_num = match.group(1)
                            # 작업명 추출
                            if style in ['dot_style', 'hash_style']:
                                work_name = match.group(2) if len(match.groups()) > 1 else ''
                            else:
                                # 같은 행 또는 다음 열에서 작업명 찾기
                                work_name = ''
                                for k in range(j+1, min(j+5, len(df.columns))):
                                    next_cell = df.iloc[i, k]
                                    if pd.notna(next_cell):
                                        work_name = str(next_cell).strip()
                                        break
                            
                            hopyo_list.append({
                                'row': i,
                                'col': j,
                                'hopyo': hopyo_num,
                                'work': work_name,
                                'style': style,
                                'full_text': cell_str
                            })
                            break
        
        # 스타일 판단
        if hopyo_list:
            main_style = hopyo_list[0]['style']
            print(f"\n구조 스타일: {main_style}")
            
            # 기존 일위대가 스타일인지 확인
            if main_style in ['je_style', 'hopyo_style']:
                print("→ 기존 일위대가 구조")
            else:
                print("→ 단가산출 구조")
        
        # 각 호표의 데이터 범위 분석
        print(f"\n총 {len(hopyo_list)}개 호표 발견:")
        
        for idx, hopyo in enumerate(hopyo_list[:5]):  # 처음 5개만
            print(f"\n[{hopyo['hopyo']}호표] {hopyo['work']} (행 {hopyo['row']})")
            
            # 다음 호표까지의 범위
            start_row = hopyo['row']
            end_row = hopyo_list[idx+1]['row'] if idx+1 < len(hopyo_list) else min(start_row + 50, len(df))
            
            # 대분류/중분류 구분
            sub_items = []
            for i in range(start_row + 1, end_row):
                for j in range(min(5, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        # 중분류 패턴 (1. 2. 3. 등)
                        sub_match = re.match(r'^(\d+)\.\s*(.+)', cell_str)
                        if sub_match and j < 3:  # 앞쪽 열에 있는 경우만
                            sub_items.append({
                                'row': i,
                                'num': sub_match.group(1),
                                'name': sub_match.group(2)
                            })
            
            print(f"  세부항목 {len(sub_items)}개:")
            for item in sub_items[:5]:  # 처음 5개만
                print(f"    {item['num']}. {item['name']} (행 {item['row']})")
            
            # 데이터 샘플
            print(f"  데이터 샘플 (행 {start_row+1} ~ {min(start_row+10, end_row)}):")
            for i in range(start_row+1, min(start_row+10, end_row)):
                row_data = []
                for j in range(min(8, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        if cell_str:
                            row_data.append(f"{j}: {cell_str[:30]}")
                if row_data:
                    print(f"    행{i}: {' | '.join(row_data[:4])}")
        
        return hopyo_list
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return []

def main():
    # 분석할 파일들
    files_to_analyze = [
        # 기존 일위대가 시트들
        ('test1.xlsx', '일위대가'),
        ('sgs.xls', '일위대가'),
        
        # 단가산출 시트들
        ('test1.xlsx', '단가산출_산근'),
        ('est.xlsx', '단가산출'),
        ('sgs.xls', '단가산출'),
        ('ebs.xls', '일위대가_산근'),  # 이름은 일위대가_산근이지만 실제는 단가산출 구조
    ]
    
    print("호표 데이터 범위 분석")
    print("="*80)
    
    results = {}
    for file_name, sheet_name in files_to_analyze:
        if os.path.exists(file_name):
            result = analyze_hopyo_data_range(file_name, sheet_name)
            results[f"{file_name}_{sheet_name}"] = result
        else:
            print(f"\n파일을 찾을 수 없음: {file_name}")
    
    # 요약
    print("\n\n" + "="*80)
    print("분석 요약")
    print("="*80)
    
    for key, hopyo_list in results.items():
        if hopyo_list:
            style = hopyo_list[0]['style']
            structure_type = "기존 일위대가" if style in ['je_style', 'hopyo_style'] else "단가산출"
            print(f"\n{key}:")
            print(f"  - 구조: {structure_type}")
            print(f"  - 스타일: {style}")
            print(f"  - 호표 수: {len(hopyo_list)}")

if __name__ == "__main__":
    main()