"""
개선된 est.xlsx 파일 파서
정확한 호표 인식과 산출근거 추출
"""
import pandas as pd
import json
import os
import re

def parse_est_improved():
    """est.xlsx 개선된 파싱"""
    print("=" * 80)
    print("est.xlsx 개선된 파서 실행")
    print("=" * 80)
    
    # result 폴더 생성
    os.makedirs('result', exist_ok=True)
    
    # 파일 읽기
    df = pd.read_excel('est.xlsx', sheet_name='일위대가', header=None)
    print(f"시트 크기: {df.shape}")
    
    # 1. 호표 찾기 (제N호표 패턴만, 컬럼 1에서만)
    hopyo_pattern = r'제\s*(\d+)\s*호표'
    hopyo_list = []
    
    for i in range(len(df)):
        # 컬럼 1에서만 찾기 (호표가 있는 위치)
        cell = df.iloc[i, 1] if df.shape[1] > 1 else None
        if pd.notna(cell):
            cell_str = str(cell).strip()
            match = re.match(hopyo_pattern, cell_str)
            if match:
                hopyo_num = match.group(1)
                
                # 작업명은 컬럼 2에서
                work_name = ''
                if df.shape[1] > 2:
                    work_cell = df.iloc[i, 2]
                    if pd.notna(work_cell):
                        work_name = str(work_cell).strip()
                
                hopyo_list.append({
                    'row': i,
                    'num': hopyo_num,
                    'work': work_name
                })
                print(f"  발견: 제{hopyo_num}호표 (행 {i+1}): {work_name}")
    
    print(f"\n총 {len(hopyo_list)}개 호표 발견")
    
    # 2. 컬럼 위치 확인 (헤더 행 2)
    # 행 2: 호표 | 품명 | 규격 | 수량 | 단위 | 합계 | 노무비
    column_info = {
        '품명': 2,  # 품명
        '규격': 3,  # 규격 
        '수량': 4,  # 수량
        '단위': 5   # 단위
    }
    
    print(f"컬럼 매핑: {column_info}")
    
    # 3. 통합 JSON 구조 생성
    result = {
        'file': 'est.xlsx',
        'sheet': '일위대가',
        'total_ilwidae_count': len(hopyo_list),
        'ilwidae_data': []
    }
    
    # 4. 각 호표 데이터 처리
    for idx, hopyo in enumerate(hopyo_list):
        start_row = hopyo['row']
        end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(df)
        
        # 타이틀 정보
        title_data = {
            '품명': hopyo['work'],
            '규격': '',
            '단위': '',
            '수량': ''
        }
        
        # 타이틀 행에서 추가 정보 추출
        if df.shape[1] > 5:
            # 규격 (컬럼 3)
            규격_cell = df.iloc[hopyo['row'], 3]
            if pd.notna(규격_cell):
                title_data['규격'] = str(규격_cell).strip()
            
            # 단위 (컬럼 5)
            단위_cell = df.iloc[hopyo['row'], 5]
            if pd.notna(단위_cell):
                title_data['단위'] = str(단위_cell).strip()
        
        # 산출근거 항목 수집
        sangul_items = []
        for row_idx in range(start_row + 1, min(start_row + 20, end_row)):
            # 빈 행은 건너뛰기
            row_empty = True
            for col in range(2, min(6, len(df.columns))):
                if pd.notna(df.iloc[row_idx, col]):
                    row_empty = False
                    break
            
            if row_empty:
                continue
            
            row_data = {
                'row_number': row_idx,
                '품명': '',
                '규격': '',
                '단위': '',
                '수량': ''
            }
            
            # 품명 (컬럼 2)
            품명_cell = df.iloc[row_idx, column_info['품명']]
            if pd.notna(품명_cell):
                value = str(품명_cell).strip()
                # 호표가 아니고 숫자만도 아닌 경우
                if value and '호표' not in value and not re.match(r'^[\d.,]+$', value):
                    row_data['품명'] = value
                else:
                    continue  # 품명이 없으면 이 행은 건너뛰기
            else:
                continue
            
            # 규격 (컬럼 3)
            if column_info['규격'] < len(df.columns):
                규격_cell = df.iloc[row_idx, column_info['규격']]
                if pd.notna(규격_cell):
                    row_data['규격'] = str(규격_cell).strip()
            
            # 수량 (컬럼 4)
            if column_info['수량'] < len(df.columns):
                수량_cell = df.iloc[row_idx, column_info['수량']]
                if pd.notna(수량_cell):
                    row_data['수량'] = str(수량_cell).strip()
            
            # 단위 (컬럼 5)
            if column_info['단위'] < len(df.columns):
                단위_cell = df.iloc[row_idx, column_info['단위']]
                if pd.notna(단위_cell):
                    row_data['단위'] = str(단위_cell).strip()
            
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
        
        if idx < 3:  # 처음 3개만 출력
            print(f"\n호표 {hopyo['num']}: {hopyo['work']}")
            print(f"  산출근거: {len(sangul_items)}개 항목")
            if sangul_items:
                print(f"  예시: {sangul_items[0]['품명']} | {sangul_items[0]['규격']} | {sangul_items[0]['단위']} | {sangul_items[0]['수량']}")
    
    # 5. JSON 파일로 저장
    output_file = 'result/est_ilwidae_improved.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ 개선된 파싱 완료: {output_file}")
    print(f"   - 일위대가 개수: {len(hopyo_list)}")
    
    # 통계 출력
    total_sangul = sum(len(item['산출근거']) for item in result['ilwidae_data'])
    print(f"   - 총 산출근거 항목: {total_sangul}개")
    
    return result

def verify_est_results():
    """결과 검증"""
    print("\n" + "=" * 80)
    print("est.xlsx 결과 검증")
    print("=" * 80)
    
    with open('result/est_ilwidae_improved.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"총 일위대가: {data['total_ilwidae_count']}개")
    
    # 필드 통계
    total_sangul = 0
    filled_counts = {'품명': 0, '규격': 0, '단위': 0, '수량': 0}
    
    for item in data['ilwidae_data']:
        for sangul in item['산출근거']:
            total_sangul += 1
            for field in filled_counts.keys():
                if sangul.get(field, ''):
                    filled_counts[field] += 1
    
    print(f"\n[필드 채움 비율]")
    print(f"총 산출근거 항목: {total_sangul}개")
    
    for field, count in filled_counts.items():
        if total_sangul > 0:
            percentage = (count / total_sangul * 100)
            status = "✅" if percentage >= 80 else "⚠️" if percentage >= 50 else "❌"
            print(f"  {status} {field}: {count}/{total_sangul} ({percentage:.1f}%)")

if __name__ == "__main__":
    # 파싱 실행
    result = parse_est_improved()
    
    # 결과 검증
    verify_est_results()