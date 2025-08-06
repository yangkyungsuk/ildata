import pandas as pd
import json
import warnings
import sys
import io
from typing import Dict, List, Any

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')

def parse_test1():
    """test1.xlsx 파일 전용 파서
    
    파일 특징:
    - 시트명: 일위대가_산근
    - 구조: 호표 | 품명 | 규격 | 단위 | 수량 형태로 각 열에 분리
    - 호표 형식: "제N호표"
    """
    
    file_path = 'test1.xlsx'
    sheet_name = '일위대가_산근'
    
    print(f"{'='*60}")
    print(f"{file_path} 파싱 시작")
    print(f"시트: {sheet_name}")
    print(f"{'='*60}")
    
    # 엑셀 읽기
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # 헤더 찾기 - test1은 1행에 주요 헤더가 있음
    header_row = 1  # 고정값
    column_map = {
        '호표': 0,
        '품명': 1,
        '규격': 2,
        '단위': 3,
        '수량': 4,
        '비고': 11  # 보통 비고는 마지막 열 근처에 위치
    }
    
    print(f"헤더 위치: 행 {header_row}")
    print(f"컬럼 매핑: {column_map}")
    
    if header_row is None:
        print("헤더를 찾을 수 없습니다!")
        return
    
    # 호표 데이터 추출
    hopyo_data = {}
    current_hopyo = None
    
    for row_idx in range(header_row + 1, len(df)):
        row = df.iloc[row_idx]
        
        # 호표 확인
        if '호표' in column_map:
            hopyo_cell = row[column_map['호표']]
            if pd.notna(hopyo_cell) and '호표' in str(hopyo_cell):
                # 호표 번호 추출
                import re
                match = re.search(r'제?\s*(\d+)\s*호표', str(hopyo_cell))
                if match:
                    hopyo_num = int(match.group(1))
                    current_hopyo = f"호표{hopyo_num}"
                    
                    # 호표 행의 정보 추출
                    work_name = ""
                    work_spec = ""
                    work_unit = ""
                    
                    if '품명' in column_map and pd.notna(row[column_map['품명']]):
                        work_name = str(row[column_map['품명']]).strip()
                    if '규격' in column_map and pd.notna(row[column_map['규격']]):
                        work_spec = str(row[column_map['규격']]).strip()
                    if '단위' in column_map and pd.notna(row[column_map['단위']]):
                        work_unit = str(row[column_map['단위']]).strip()
                    
                    hopyo_data[current_hopyo] = {
                        "호표번호": hopyo_num,
                        "작업명": work_name,
                        "규격": work_spec,
                        "단위": work_unit,
                        "세부항목": []
                    }
                    
                    print(f"\n호표{hopyo_num} 발견:")
                    print(f"  작업명: {work_name}")
                    print(f"  규격: {work_spec}")
                    print(f"  단위: {work_unit}")
                    continue
        
        # 세부 항목 추출
        if current_hopyo and '품명' in column_map:
            품명_cell = row[column_map['품명']]
            
            if pd.notna(품명_cell):
                품명 = str(품명_cell).strip()
                
                # 요약 행 스킵
                if any(skip in 품명 for skip in ['합계', '소계', '재료비', '노무비', '경비']):
                    continue
                
                # 빈 행 스킵
                if not 품명:
                    continue
                
                # 세부 항목 데이터 수집
                item = {"품명": 품명}
                
                if '규격' in column_map and pd.notna(row[column_map['규격']]):
                    item['규격'] = str(row[column_map['규격']]).strip()
                else:
                    item['규격'] = ""
                
                if '단위' in column_map and pd.notna(row[column_map['단위']]):
                    item['단위'] = str(row[column_map['단위']]).strip()
                else:
                    item['단위'] = ""
                
                if '수량' in column_map and pd.notna(row[column_map['수량']]):
                    try:
                        item['수량'] = float(str(row[column_map['수량']]).replace(',', ''))
                    except:
                        item['수량'] = 0
                else:
                    item['수량'] = 0
                
                if '비고' in column_map and pd.notna(row[column_map['비고']]):
                    item['비고'] = str(row[column_map['비고']]).strip()
                else:
                    item['비고'] = ""
                
                hopyo_data[current_hopyo]['세부항목'].append(item)
    
    # 결과 정리
    for hopyo_key, data in hopyo_data.items():
        data['항목수'] = len(data['세부항목'])
        print(f"\n{hopyo_key}: {data['항목수']}개 항목")
    
    # 결과 저장
    result = {
        "file": file_path,
        "sheet": sheet_name,
        "hopyo_count": len(hopyo_data),
        "hopyo_data": hopyo_data
    }
    
    output_file = "test1_parsed.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'='*60}")
    print(f"파싱 완료!")
    print(f"저장 위치: {output_file}")
    print(f"총 호표 수: {len(hopyo_data)}")
    print(f"총 세부항목 수: {sum(data['항목수'] for data in hopyo_data.values())}")
    print(f"{'='*60}")

if __name__ == "__main__":
    parse_test1()