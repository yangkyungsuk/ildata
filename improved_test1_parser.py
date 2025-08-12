"""
개선된 test1.xlsx 일위대가 파서
- 중복 호표 필터링
- 규격/단위 정확한 추출
- 빈 작업명 호표 제외
"""
import pandas as pd
import json
import os
import re

# result 폴더 생성
os.makedirs('result', exist_ok=True)

def parse_test1_improved():
    """개선된 test1.xlsx 파싱"""
    
    print("=" * 70)
    print("개선된 test1.xlsx 파서 실행")
    print("=" * 70)
    
    # 1. 엑셀 파일 읽기
    df = pd.read_excel('test1.xlsx', sheet_name='일위대가_산근', header=None)
    print(f"시트 크기: {df.shape}")
    
    # 2. 호표 찾기 (중복 제거, 빈 작업명 제외)
    hopyo_pattern = r'제\s*(\d+)\s*호표'
    hopyo_list = []
    seen_nums = set()  # 이미 처리한 호표 번호 추적
    
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
                            if next_str and not next_str.replace('.', '').replace(',', '').isdigit():
                                work_name = next_str
                                break
                    
                    # 중복 번호이거나 작업명이 없으면 건너뛰기
                    if hopyo_num in seen_nums or not work_name:
                        print(f"  건너뜀: 제{hopyo_num}호표 (행 {i+1}) - {'중복' if hopyo_num in seen_nums else '빈 작업명'}")
                        break
                    
                    seen_nums.add(hopyo_num)
                    hopyo_list.append({
                        'row': i,
                        'num': hopyo_num,
                        'work': work_name
                    })
                    print(f"  발견: 제{hopyo_num}호표 (행 {i+1}): {work_name}")
                    break
    
    print(f"\n총 {len(hopyo_list)}개 유효한 호표 발견")
    
    # 3. 컬럼 위치 자동 감지 (개선된 버전)
    def detect_column_positions(df, hopyo_list):
        """산출근거 항목들의 컬럼 위치 자동 감지"""
        
        # 첫 번째 호표의 산출근거 영역에서 패턴 분석
        if len(hopyo_list) == 0:
            return None
        
        first_hopyo = hopyo_list[0]
        start_row = first_hopyo['row'] + 1
        end_row = start_row + 5  # 처음 5개 행만 분석
        
        column_info = {
            '품명': None,
            '규격': None,
            '단위': None,
            '수량': None
        }
        
        # 각 행을 분석하여 컬럼 패턴 파악
        for row_idx in range(start_row, min(end_row, len(df))):
            row_data = []
            for col_idx in range(min(10, len(df.columns))):
                cell = df.iloc[row_idx, col_idx]
                if pd.notna(cell):
                    row_data.append((col_idx, str(cell).strip()))
                else:
                    row_data.append((col_idx, ''))
            
            # 패턴 분석
            for idx, (col_idx, value) in enumerate(row_data):
                if value:
                    # 품명: 첫 번째 텍스트 (철근공, 보통인부 등)
                    if column_info['품명'] is None and re.match(r'^[가-힣]+', value):
                        column_info['품명'] = col_idx
                    
                    # 단위: 인, %, M3 등
                    elif value in ['인', '%', 'M3', 'M2', 'M', 'KG', 'TON', '개', '대', 'L', 'EA', 'HR', '조', 'm3']:
                        column_info['단위'] = col_idx
                    
                    # 수량: 소수점 숫자
                    elif re.match(r'^\d+\.?\d*$', value) and float(value) < 100:
                        if column_info['수량'] is None:
                            column_info['수량'] = col_idx
        
        # 규격은 품명과 단위 사이로 추정
        if column_info['품명'] is not None and column_info['단위'] is not None:
            column_info['규격'] = column_info['품명'] + 1
        
        print(f"\n컬럼 위치 감지 결과:")
        for key, val in column_info.items():
            print(f"  {key}: 컬럼 {val}")
        
        return column_info
    
    column_info = detect_column_positions(df, hopyo_list)
    
    # 4. 통합 JSON 구조 생성
    result = {
        'file': 'test1.xlsx',
        'sheet': '일위대가_산근',
        'total_ilwidae_count': len(hopyo_list),
        'validation': {
            'duplicate_removed': len(seen_nums),
            'empty_work_skipped': True
        },
        'ilwidae_data': []
    }
    
    # 5. 각 호표 데이터 처리
    for idx, hopyo in enumerate(hopyo_list):
        start_row = hopyo['row']
        end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(df)
        
        # 타이틀 정보 추출
        title_data = {
            '품명': hopyo['work'],
            '규격': '',
            '단위': '',
            '수량': ''
        }
        
        # 호표 행에서 추가 정보 추출
        for col_idx in range(len(df.columns)):
            cell = df.iloc[hopyo['row'], col_idx]
            if pd.notna(cell):
                value = str(cell).strip()
                # 단위 찾기
                if value in ['TON', 'M3', 'M2', 'M', '개', '대']:
                    title_data['단위'] = value
                # 수량 찾기 (큰 숫자)
                elif re.match(r'^\d+\.?\d*$', value) and float(value) >= 1:
                    if not title_data['수량']:
                        title_data['수량'] = value
        
        # 산출근거 항목 수집
        sangul_items = []
        for row_idx in range(start_row + 1, min(start_row + 20, end_row)):
            row_data = {
                'row_number': row_idx,
                '품명': '',
                '규격': '',
                '단위': '',
                '수량': '',
                '비고': ''
            }
            
            # 각 컬럼에서 데이터 추출
            has_content = False
            
            # 품명 추출
            if column_info['품명'] is not None:
                cell = df.iloc[row_idx, column_info['품명']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data['품명'] = value
                        has_content = True
            
            # 규격 추출
            if column_info['규격'] is not None:
                cell = df.iloc[row_idx, column_info['규격']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data['규격'] = value
            
            # 단위 추출
            if column_info['단위'] is not None:
                cell = df.iloc[row_idx, column_info['단위']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data['단위'] = value
            
            # 수량 추출
            if column_info['수량'] is not None:
                cell = df.iloc[row_idx, column_info['수량']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data['수량'] = value
            
            # 내용이 있는 행만 추가
            if has_content:
                sangul_items.append(row_data)
        
        ilwidae_item = {
            'ilwidae_no': hopyo['num'],
            'ilwidae_title': title_data,
            'position': {
                'start_row': start_row,
                'end_row': end_row,
                'total_rows': end_row - start_row
            },
            '산출근거': sangul_items
        }
        
        result['ilwidae_data'].append(ilwidae_item)
        
        print(f"\n호표 {hopyo['num']}: {hopyo['work']}")
        print(f"  산출근거: {len(sangul_items)}개 항목")
        if sangul_items:
            print(f"  예시: {sangul_items[0]['품명']} ({sangul_items[0]['규격']}) - {sangul_items[0]['단위']} {sangul_items[0]['수량']}")
    
    # 6. JSON 파일로 저장
    output_file = 'result/test1_ilwidae_improved.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ 개선된 파싱 완료: {output_file}")
    print(f"   - 일위대가 개수: {len(hopyo_list)}")
    print(f"   - 중복 제거됨")
    print(f"   - 규격/단위 추출 개선")
    
    return result

def verify_improved_results():
    """개선된 결과 검증"""
    print("\n" + "=" * 70)
    print("개선된 결과 검증")
    print("=" * 70)
    
    # JSON 파일 읽기
    with open('result/test1_ilwidae_improved.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"\n총 일위대가: {data['total_ilwidae_count']}개")
    
    # 규격/단위 필드 통계
    total_sangul = 0
    filled_counts = {
        '품명': 0,
        '규격': 0,
        '단위': 0,
        '수량': 0
    }
    
    for item in data['ilwidae_data']:
        for sangul in item['산출근거']:
            total_sangul += 1
            for field in filled_counts.keys():
                if sangul.get(field, ''):
                    filled_counts[field] += 1
    
    print(f"\n[필드 채움 비율]")
    print(f"총 산출근거 항목: {total_sangul}개")
    for field, count in filled_counts.items():
        percentage = (count / total_sangul * 100) if total_sangul > 0 else 0
        status = "✅" if percentage > 80 else "⚠️" if percentage > 50 else "❌"
        print(f"  {status} {field}: {count}/{total_sangul} ({percentage:.1f}%)")
    
    # 샘플 데이터 출력
    print(f"\n[샘플 데이터]")
    if data['ilwidae_data']:
        first = data['ilwidae_data'][0]
        print(f"제{first['ilwidae_no']}호표: {first['ilwidae_title']['품명']}")
        for i, sangul in enumerate(first['산출근거'][:3], 1):
            print(f"  {i}. {sangul['품명']} | {sangul['규격']} | {sangul['단위']} | {sangul['수량']}")

if __name__ == "__main__":
    # 개선된 파서 실행
    result = parse_test1_improved()
    
    # 결과 검증
    verify_improved_results()