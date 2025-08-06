import pandas as pd
import json
import warnings
import sys
import io
from typing import Dict, List, Any
import re

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')

def parse_est():
    """est.xlsx 파일 전용 파서
    
    파일 특징:
    - 시트명: 일위대가 (일위대가총괄표가 아님!)
    - 구조: 호표 | 품명 | 규격 | 수량 | 단위 형태
    - 호표 형식: "제N호표"
    - 특이사항: 세부항목이 들여쓰기 되어 있음 (첫 열이 비어있음)
    """
    
    file_path = 'est.xlsx'
    sheet_name = '일위대가'
    
    print(f"{'='*60}")
    print(f"{file_path} 파싱 시작")
    print(f"시트: {sheet_name}")
    print(f"{'='*60}")
    
    # 엑셀 읽기
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # 호표 데이터 추출
    hopyo_data = {}
    current_hopyo = None
    
    # 호표 패턴
    hopyo_pattern = re.compile(r'제\s*(\d+)\s*호표')
    
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        
        # 호표 찾기 (보통 1열에 위치)
        for col_idx in range(min(3, len(row))):
            if pd.notna(row[col_idx]):
                cell_str = str(row[col_idx])
                hopyo_match = hopyo_pattern.search(cell_str)
                
                if hopyo_match:
                    hopyo_num = int(hopyo_match.group(1))
                    current_hopyo = f"호표{hopyo_num}"
                    
                    # 호표 행의 정보 추출
                    work_name = ""
                    work_spec = ""
                    work_unit = ""
                    
                    # 호표 다음 열부터 확인
                    if col_idx + 1 < len(row) and pd.notna(row[col_idx + 1]):
                        work_name = str(row[col_idx + 1]).strip()
                    if col_idx + 2 < len(row) and pd.notna(row[col_idx + 2]):
                        work_spec = str(row[col_idx + 2]).strip()
                    
                    # 단위는 보통 수량 다음 열에 위치 (4번째 또는 5번째 열)
                    if col_idx + 4 < len(row) and pd.notna(row[col_idx + 4]):
                        work_unit = str(row[col_idx + 4]).strip()
                    elif col_idx + 3 < len(row) and pd.notna(row[col_idx + 3]):
                        # 수량이 비어있고 단위가 3번째에 있는 경우
                        val = str(row[col_idx + 3]).strip()
                        if not val.replace('.', '').replace(',', '').isdigit():
                            work_unit = val
                    
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
                    break
        
        # 세부 항목 추출
        if current_hopyo:
            # 1열(인덱스 1)에 호표가 없고, 2열(인덱스 2)에 품명이 있으면 세부 항목
            col1_val = str(row[1]).strip() if pd.notna(row[1]) else ''
            col2_val = str(row[2]).strip() if pd.notna(row[2]) else ''
            
            # 1열에 호표가 없고 2열에 품명이 있으면 세부 항목
            if '호표' not in col1_val and col2_val:
                품명 = col2_val
                
                # 빈 값이나 숫자만 있는 경우 스킵
                if not 품명 or 품명.replace('.', '').replace(',', '').isdigit():
                    continue
                
                # 0만 있는 행 스킵
                all_zero = True
                for col in row[2:]:
                    if pd.notna(col) and str(col).strip() != '0':
                        all_zero = False
                        break
                if all_zero and '0' in str(row[2:]).strip():
                    continue
                
                # 세부 항목 데이터 수집
                item = {"품명": 품명}
                
                # 규격 (3번째 열, 인덱스 3)
                if pd.notna(row[3]):
                    item['규격'] = str(row[3]).strip()
                else:
                    item['규격'] = ""
                
                # 수량 (4번째 열, 인덱스 4)
                if pd.notna(row[4]):
                    try:
                        item['수량'] = float(str(row[4]).replace(',', ''))
                    except:
                        item['수량'] = 0
                else:
                    item['수량'] = 0
                
                # 단위 (5번째 열, 인덱스 5)
                if pd.notna(row[5]):
                    item['단위'] = str(row[5]).strip()
                else:
                    item['단위'] = ""
                
                # 비고 (보통 7번째 이후 열에 있을 수 있음)
                item['비고'] = ""
                for col_idx in range(7, len(row)):
                    if pd.notna(row[col_idx]):
                        val = str(row[col_idx]).strip()
                        if val and not val.replace('.', '').replace(',', '').isdigit():
                            item['비고'] = val
                            break
                
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
    
    output_file = "est_parsed.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'='*60}")
    print(f"파싱 완료!")
    print(f"저장 위치: {output_file}")
    print(f"총 호표 수: {len(hopyo_data)}")
    print(f"총 세부항목 수: {sum(data['항목수'] for data in hopyo_data.values())}")
    print(f"{'='*60}")

if __name__ == "__main__":
    parse_est()