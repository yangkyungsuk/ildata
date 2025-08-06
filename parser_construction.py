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

def parse_construction():
    """건축구조내역2.xlsx 파일 전용 파서
    
    파일 특징:
    - 시트명: 일위대가
    - 구조: "작업명 규격 단위 (호표 N)" 형식으로 한 셀에 포함
    - 특이사항: 코드 문자열이 포함되어 있음
    - 세부항목이 첫 번째 열에 위치
    """
    
    file_path = '건축구조내역2.xlsx'
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
    hopyo_pattern = re.compile(r'\(\s*호표\s*(\d+)\s*\)')
    
    # 일반적인 단위들
    units = ['M', 'm', 'EA', 'ea', '개소', '식', 'TON', '㎥', '㎡', '㎢', 
             '개', '본', '인', '시간', '일', 'KG', 'kg', '대', 'SET', 'set',
             'ℓ', 'L', '회', '㎞', 'km', 'HR', 'hr', '조', '주']
    
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        
        # 첫 번째 열에서 호표 찾기
        if pd.notna(row[0]):
            cell_str = str(row[0])
            
            # 호표 찾기
            hopyo_match = hopyo_pattern.search(cell_str)
            if hopyo_match:
                hopyo_num = int(hopyo_match.group(1))
                current_hopyo = f"호표{hopyo_num}"
                
                # 호표 정보 파싱
                work_name = ""
                work_spec = ""
                work_unit = ""
                
                # 호표 부분 제거
                before_hopyo = cell_str[:hopyo_match.start()].strip()
                
                # 첫 부분에 코드가 있으면 제거 (예: "M-101 정밀여과장치 설치")
                if ' ' in before_hopyo:
                    parts = before_hopyo.split(' ', 1)
                    # 첫 부분이 코드처럼 보이면 (대문자+숫자 조합)
                    if len(parts[0]) < 10 and any(c.isdigit() for c in parts[0]):
                        work_name = parts[1] if len(parts) > 1 else before_hopyo
                    else:
                        work_name = before_hopyo
                else:
                    work_name = before_hopyo
                
                # 더블 스페이스로 구분된 경우
                if '  ' in work_name:
                    parts = [p.strip() for p in work_name.split('  ') if p.strip()]
                    if parts:
                        work_name = parts[0]
                        if len(parts) > 1:
                            # 나머지에서 단위 찾기
                            remaining = ' '.join(parts[1:])
                            words = remaining.split()
                            
                            # 뒤에서부터 단위 찾기
                            for i in range(len(words)-1, -1, -1):
                                if words[i] in units:
                                    work_unit = words[i]
                                    work_spec = ' '.join(words[:i]).strip()
                                    break
                            
                            if not work_unit and remaining:
                                work_spec = remaining
                
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
        if current_hopyo and pd.notna(row[0]):
            품명 = str(row[0]).strip()
            
            # 호표가 아니고, 합계행이 아니면 세부 항목
            if not hopyo_pattern.search(품명) and '[' not in 품명:
                # 이상한 코드 문자열 스킵 (길이가 너무 길거나 특정 패턴)
                if len(품명) > 50 or (len(품명) > 20 and 품명.replace('0','').replace('1','').replace('2','').replace('3','').replace('4','').replace('5','').replace('6','').replace('7','').replace('8','').replace('9','').replace('A','').replace('B','').replace('C','').replace('D','').replace('E','').replace('F','') == ''):
                    continue
                
                # 세부 항목 데이터 수집
                item = {"품명": 품명}
                
                # 규격 (두 번째 열)
                if pd.notna(row[1]):
                    item['규격'] = str(row[1]).strip()
                else:
                    item['규격'] = ""
                
                # 단위 (세 번째 열)
                if pd.notna(row[2]):
                    item['단위'] = str(row[2]).strip()
                else:
                    item['단위'] = ""
                
                # 수량 (네 번째 열)
                if pd.notna(row[3]):
                    try:
                        item['수량'] = float(str(row[3]).replace(',', ''))
                    except:
                        item['수량'] = 0
                else:
                    item['수량'] = 0
                
                # 비고 (마지막 열이나 특정 위치)
                # 건축구조 파일은 비고란이 명확하지 않으므로 빈 값으로 설정
                item['비고'] = ""
                
                # 데이터가 있는 항목만 추가
                if item['품명'] and (item['수량'] > 0 or item['단위'] or item['규격']):
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
    
    output_file = "construction_parsed.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'='*60}")
    print(f"파싱 완료!")
    print(f"저장 위치: {output_file}")
    print(f"총 호표 수: {len(hopyo_data)}")
    print(f"총 세부항목 수: {sum(data['항목수'] for data in hopyo_data.values())}")
    print(f"{'='*60}")

if __name__ == "__main__":
    parse_construction()