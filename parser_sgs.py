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

def parse_sgs():
    """sgs.xls 파일 전용 파서
    
    파일 특징:
    - 시트명: 일위대가
    - 구조: 호표가 콜론(:) 형식으로 한 셀에 포함
    - 호표 형식: "제 N 호표 ：작업명 (규격) 단위 당"
    - 특이사항: 계산 행(계 ÷ 120, 계 x 0.5 등)도 포함
    """
    
    file_path = 'sgs.xls'
    sheet_name = '일위대가'
    
    print(f"{'='*60}")
    print(f"{file_path} 파싱 시작")
    print(f"시트: {sheet_name}")
    print(f"{'='*60}")
    
    # 엑셀 읽기 (xls 파일이므로 xlrd 엔진 사용)
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
    
    # 헤더 찾기 - sgs는 2행에 헤더가 있음
    header_row = 2  # 고정값
    column_map = {
        '공종': 0,  # sgs는 품명 대신 공종 사용
        '규격': 1,
        '수량': 2,
        '단위': 3,
        '비고': 11  # 마지막 열 근처
    }
    
    print(f"헤더 위치: 행 {header_row}")
    print(f"컬럼 매핑: {column_map}")
    
    if header_row is None:
        print("헤더를 찾을 수 없습니다!")
        return
    
    # 호표 데이터 추출
    hopyo_data = {}
    current_hopyo = None
    
    # 호표 패턴
    hopyo_pattern = re.compile(r'제\s*(\d+)\s*호표')
    
    for row_idx in range(header_row + 1, len(df)):
        row = df.iloc[row_idx]
        
        # 첫 번째 셀에서 호표 찾기
        first_cell = row[0] if pd.notna(row[0]) else ""
        first_cell_str = str(first_cell)
        
        # 호표 찾기
        hopyo_match = hopyo_pattern.search(first_cell_str)
        if hopyo_match:
            hopyo_num = int(hopyo_match.group(1))
            current_hopyo = f"호표{hopyo_num}"
            
            # 호표 정보 파싱
            work_name = ""
            work_spec = ""
            work_unit = ""
            
            # 콜론으로 분리
            if '：' in first_cell_str:
                parts = first_cell_str.split('：', 1)
                if len(parts) > 1:
                    work_info = parts[1].strip()
                    
                    # "당" 제거
                    if work_info.endswith(' 당'):
                        work_info = work_info[:-2].strip()
                    
                    # 괄호로 규격 분리
                    paren_match = re.search(r'\((.*?)\)', work_info)
                    if paren_match:
                        work_spec = paren_match.group(1)
                        work_name = work_info[:paren_match.start()].strip()
                        work_unit = work_info[paren_match.end():].strip()
                    else:
                        # 마지막 단어를 단위로 추정
                        words = work_info.split()
                        if words:
                            # 일반적인 단위들
                            units = ['M', 'm', 'EA', '개소', '㎡', '인', '기', '본', '대']
                            last_word = words[-1]
                            
                            if last_word in units:
                                work_unit = last_word
                                work_name = ' '.join(words[:-1])
                            else:
                                work_name = work_info
            
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
        if current_hopyo and '공종' in column_map:
            공종_cell = row[column_map['공종']]
            
            if pd.notna(공종_cell):
                공종 = str(공종_cell).strip()
                
                # 빈 행 스킵
                if not 공종:
                    continue
                
                # 최종 합계 행 확인 (많은 공백이 있는 '계')
                if '계' in 공종 and len(공종) > 10:
                    # 이것은 호표의 마지막 합계 행
                    continue
                
                # 세부 항목 데이터 수집
                item = {"품명": 공종}  # JSON에서는 품명으로 통일
                
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
    
    output_file = "sgs_parsed.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'='*60}")
    print(f"파싱 완료!")
    print(f"저장 위치: {output_file}")
    print(f"총 호표 수: {len(hopyo_data)}")
    print(f"총 세부항목 수: {sum(data['항목수'] for data in hopyo_data.values())}")
    print(f"{'='*60}")

if __name__ == "__main__":
    parse_sgs()