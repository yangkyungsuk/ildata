"""
test1.xlsx 단가산출_산근 시트 파서 (목록 검증 기능 포함)
스타일: dot_style (1.작업명 형식)
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def read_sangcul_list(file_path):
    """단가산출목록표 읽기"""
    try:
        # 목록 시트 이름 찾기
        xls = pd.ExcelFile(file_path)
        list_sheet = None
        for sheet in xls.sheet_names:
            if '단가산출' in sheet and '목록' in sheet:
                list_sheet = sheet
                break
        
        if not list_sheet:
            return None
        
        # 목록 읽기
        df = pd.read_excel(file_path, sheet_name=list_sheet, header=None)
        
        # 호표 정보 추출
        hopyo_info = []
        for i in range(len(df)):
            cell = df.iloc[i, 0] if 0 < len(df.columns) else None
            if pd.notna(cell):
                cell_str = str(cell).strip()
                # 제N호표 패턴
                match = re.match(r'제(\d+)호표', cell_str)
                if match:
                    hopyo_num = match.group(1)
                    # 같은 행에서 품명, 규격 찾기
                    work_name = df.iloc[i, 1] if 1 < len(df.columns) else ''
                    spec = df.iloc[i, 2] if 2 < len(df.columns) else ''
                    
                    hopyo_info.append({
                        'num': hopyo_num,
                        'work': str(work_name).strip() if pd.notna(work_name) else '',
                        'spec': str(spec).strip() if pd.notna(spec) else ''
                    })
        
        print(f"\n[단가산출목록표 정보]")
        print(f"총 {len(hopyo_info)}개 호표 발견")
        for info in hopyo_info:
            print(f"  제{info['num']}호표: {info['work']} ({info['spec']})")
        
        return hopyo_info
        
    except Exception as e:
        print(f"목록 읽기 오류: {str(e)}")
        return None

def parse_test1_sangcul():
    """test1.xlsx의 단가산출_산근 시트 파싱"""
    
    file_path = 'test1.xlsx'
    sheet_name = '단가산출_산근'
    
    try:
        # 먼저 목록 정보 읽기
        list_info = read_sangcul_list(file_path)
        
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        print(f"\n시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
        # 헤더 찾기 (산출근거가 있는 행)
        header_row = None
        sangcul_col = 0  # 산출근거는 보통 0열
        
        for i in range(min(10, len(df))):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell) and '산출근거' in str(cell):
                    header_row = i
                    sangcul_col = j
                    print(f"헤더 행 발견: {header_row}")
                    break
            if header_row is not None:
                break
        
        # 호표 패턴: 숫자.작업명 형식
        hopyo_pattern = r'^(\d+)\.\s*(.+)'
        
        # 모든 호표 찾기
        hopyo_list = []
        for i in range(header_row + 1 if header_row else 0, len(df)):
            cell = df.iloc[i, sangcul_col]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(hopyo_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    work_name = match.group(2).strip()
                    
                    # 대분류 호표인지 확인 (1, 2, 3 등 연속된 번호)
                    if len(hopyo_list) == 0 or int(hopyo_num) == len(hopyo_list) + 1:
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name,
                            'full_text': cell_str
                        })
        
        print(f"\n[파싱 결과]")
        print(f"총 {len(hopyo_list)}개 호표 발견")
        
        # 목록과 비교
        if list_info:
            print(f"\n[검증 결과]")
            print(f"목록표: {len(list_info)}개, 실제 파싱: {len(hopyo_list)}개")
            
            if len(list_info) == len(hopyo_list):
                print("✓ 호표 개수 일치")
            else:
                print(f"✗ 호표 개수 불일치")
            
            # 각 호표별 비교
            for i in range(min(len(list_info), len(hopyo_list))):
                list_item = list_info[i]
                parsed_item = hopyo_list[i]
                
                if list_item['num'] == parsed_item['num']:
                    # 작업명 비교 (일부만 일치해도 OK)
                    if list_item['work'] in parsed_item['work'] or parsed_item['work'] in list_item['work']:
                        print(f"✓ 제{list_item['num']}호표: {list_item['work']} = {parsed_item['work'][:30]}...")
                    else:
                        print(f"? 제{list_item['num']}호표: 목록({list_item['work']}) ≠ 파싱({parsed_item['work'][:30]}...)")
                else:
                    print(f"✗ 호표 번호 불일치: 목록(제{list_item['num']}호표) ≠ 파싱({parsed_item['num']})")
        
        # 각 호표의 데이터 추출
        result = {
            'file': file_path,
            'sheet': sheet_name,
            'hopyo_count': len(hopyo_list),
            'hopyo_data': {}
        }
        
        for idx, hopyo in enumerate(hopyo_list):
            hopyo_key = f"호표{hopyo['num']}"
            
            # 다음 호표까지의 범위
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(df)
            
            # 산출근거 내용 수집
            sangcul_content = []
            
            for i in range(start_row, end_row):
                cell = df.iloc[i, sangcul_col]
                if pd.notna(cell):
                    content = str(cell).strip()
                    if content:
                        # 중분류 패턴 확인 (숫자. 형식)
                        sub_match = re.match(r'^(\d+)\.\s*(.+)', content)
                        if sub_match:
                            sub_num = sub_match.group(1)
                            sub_name = sub_match.group(2)
                            
                            # 대분류가 아닌 중분류인지 확인
                            if int(sub_num) < 100:  # 임의의 기준
                                sangcul_content.append({
                                    'row': i,
                                    'type': 'sub_item',
                                    'num': sub_num,
                                    'name': sub_name,
                                    'content': content
                                })
                        else:
                            # 일반 내용
                            sangcul_content.append({
                                'row': i,
                                'type': 'content',
                                'content': content
                            })
            
            result['hopyo_data'][hopyo_key] = {
                '호표번호': hopyo['num'],
                '작업명': hopyo['work'],
                '시작행': start_row,
                '종료행': end_row,
                '산출근거': sangcul_content,
                '항목수': len(sangcul_content)
            }
        
        # JSON 파일로 저장
        output_file = 'test1_sangcul_parsed_v2.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n파싱 완료. 결과 저장: {output_file}")
        return result
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    parse_test1_sangcul()