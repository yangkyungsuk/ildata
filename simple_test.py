"""
간단한 테스트 - test1.xlsx 파일 처리
"""
import pandas as pd
import json
import os
import re

# result 폴더 생성
os.makedirs('result', exist_ok=True)

# test1.xlsx 읽기
print("test1.xlsx 파일 처리 시작...")

try:
    # 일위대가_산근 시트 읽기
    df = pd.read_excel('test1.xlsx', sheet_name='일위대가_산근', header=None)
    print(f"일위대가_산근 시트 크기: {df.shape}")
    
    # 호표 찾기 (제N호표 패턴)
    hopyo_list = []
    hopyo_pattern = r'제\s*(\d+)\s*호표'
    
    for i in range(len(df)):
        for j in range(min(5, len(df.columns))):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(hopyo_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    
                    # 작업명 찾기
                    work_name = ''
                    for k in range(j+1, min(j+5, len(df.columns))):
                        next_cell = df.iloc[i, k]
                        if pd.notna(next_cell):
                            next_str = str(next_cell).strip()
                            if next_str and not next_str.isdigit():
                                work_name = next_str
                                break
                    
                    hopyo_list.append({
                        'row': i,
                        'num': hopyo_num,
                        'work': work_name
                    })
                    print(f"호표 발견: 제{hopyo_num}호표 - {work_name} (행 {i})")
                    break
    
    print(f"\n총 {len(hopyo_list)}개 호표 발견")
    
    # 통합 JSON 구조 생성
    result = {
        'file': 'test1.xlsx',
        'sheet': '일위대가_산근',
        'total_ilwidae_count': len(hopyo_list),
        'ilwidae_data': []
    }
    
    # 각 호표 데이터 처리
    for idx, hopyo in enumerate(hopyo_list):
        start_row = hopyo['row']
        end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(df)
        
        # 산출근거 항목 수집
        sangul_items = []
        for row_idx in range(start_row + 1, min(start_row + 20, end_row)):  # 최대 20행만
            row_data = {
                'row_number': row_idx,
                '품명': '',
                '규격': '',
                '단위': '',
                '수량': ''
            }
            
            # 각 컬럼 데이터 수집
            has_content = False
            for col_idx in range(min(10, len(df.columns))):
                cell = df.iloc[row_idx, col_idx]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        has_content = True
                        # 첫 번째 유효한 값을 품명으로
                        if col_idx == 0 or (not row_data['품명'] and not value.replace('.', '').isdigit()):
                            row_data['품명'] = value
                        # 단위 패턴
                        elif value in ['M3', 'M2', 'M', 'KG', 'TON', '개', '대', 'L', 'EA', 'M³', 'M²']:
                            row_data['단위'] = value
                        # 숫자는 수량으로
                        elif re.match(r'^\d+\.?\d*$', value):
                            if not row_data['수량']:
                                row_data['수량'] = value
            
            if has_content:
                sangul_items.append(row_data)
        
        ilwidae_item = {
            'ilwidae_no': hopyo['num'],
            'ilwidae_title': {
                '품명': hopyo['work'],
                '규격': '',
                '단위': '',
                '수량': ''
            },
            'position': {
                'start_row': start_row,
                'end_row': end_row,
                'total_rows': end_row - start_row
            },
            '산출근거': sangul_items[:10]  # 최대 10개만
        }
        
        result['ilwidae_data'].append(ilwidae_item)
    
    # JSON 파일로 저장
    output_file = 'result/test1_일위대가_산근_unified.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ 저장 완료: {output_file}")
    print(f"   - 일위대가 개수: {len(hopyo_list)}")
    print(f"   - 파일 크기: {os.path.getsize(output_file)} bytes")
    
except Exception as e:
    print(f"오류 발생: {str(e)}")
    import traceback
    traceback.print_exc()

# 다른 파일들도 간단히 확인
print("\n" + "=" * 50)
print("다른 파일 확인")

files_to_check = ['est.xlsx', 'sgs.xls', '건축구조내역.xlsx', '건축구조내역2.xlsx']

for file_name in files_to_check:
    if os.path.exists(file_name):
        try:
            xl = pd.ExcelFile(file_name)
            ilwidae_sheets = [s for s in xl.sheet_names if '일위대가' in s and '목록' not in s and '총괄' not in s]
            if ilwidae_sheets:
                print(f"{file_name}: {ilwidae_sheets[0]} 시트 발견")
                
                # 첫 번째 일위대가 시트 처리
                df = pd.read_excel(file_name, sheet_name=ilwidae_sheets[0], header=None)
                
                # 간단한 구조로 저장
                result = {
                    'file': file_name,
                    'sheet': ilwidae_sheets[0],
                    'rows': len(df),
                    'columns': len(df.columns)
                }
                
                output_file = f"result/{file_name.split('.')[0]}_{ilwidae_sheets[0]}_info.json"
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                
                print(f"  → {output_file} 저장")
                
        except Exception as e:
            print(f"{file_name}: 오류 - {str(e)}")

print("\n완료!")