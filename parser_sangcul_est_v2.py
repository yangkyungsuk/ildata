"""
est.xlsx 단가산출 시트 파서 (목록 검증 기능 포함)
스타일: dot_style (1.작업명 형식)
특징: 헤더가 1열부터 시작, 단가산출총괄표 사용
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def read_sangcul_list(file_path):
    """단가산출총괄표 읽기"""
    try:
        # 목록 시트 이름 찾기
        xls = pd.ExcelFile(file_path)
        list_sheet = None
        for sheet in xls.sheet_names:
            if '단가산출' in sheet and '총괄' in sheet:
                list_sheet = sheet
                break
        
        if not list_sheet:
            return None
        
        # 목록 읽기
        df = pd.read_excel(file_path, sheet_name=list_sheet, header=None)
        
        # 호표 정보 추출
        hopyo_info = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    # 호표 패턴: 숫자만 있고 다음 열에 작업명
                    if cell_str.isdigit():
                        num = int(cell_str)
                        if 1 <= num <= 100:  # 호표 번호 범위
                            # 같은 행에서 작업명 찾기
                            work_name = None
                            for k in range(j+1, len(df.columns)):
                                next_cell = df.iloc[i, k]
                                if pd.notna(next_cell):
                                    next_str = str(next_cell).strip()
                                    if next_str and not next_str.isdigit():
                                        work_name = next_str
                                        break
                            
                            if work_name:
                                hopyo_info.append({
                                    'num': str(num),
                                    'work': work_name
                                })
                                break
        
        print(f"\n[단가산출총괄표 정보]")
        print(f"총 {len(hopyo_info)}개 호표 발견")
        for info in hopyo_info:
            print(f"  {info['num']}호표: {info['work']}")
        
        return hopyo_info
        
    except Exception as e:
        print(f"목록 읽기 오류: {str(e)}")
        return None

def parse_est_sangcul():
    """est.xlsx의 단가산출 시트 파싱"""
    
    file_path = 'est.xlsx'
    sheet_name = '단가산출'
    
    try:
        # 먼저 목록 정보 읽기
        list_info = read_sangcul_list(file_path)
        
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        print(f"\n시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
        # 헤더 찾기 (산출근거가 있는 행)
        header_row = None
        sangcul_col = 1  # est는 1열에 주요 내용이 있음
        
        for i in range(min(10, len(df))):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if pd.notna(cell) and '산' in str(cell) and '근' in str(cell):
                    header_row = i
                    print(f"헤더 행 발견: {header_row}")
                    break
            if header_row is not None:
                break
        
        # 호표 패턴: 숫자.작업명 형식
        hopyo_pattern = r'^(\d+)\.\s*(.+)'
        
        # 모든 호표 찾기 (대분류만)
        hopyo_list = []
        expected_num = 1
        
        for i in range(header_row + 1 if header_row else 0, len(df)):
            cell = df.iloc[i, sangcul_col]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(hopyo_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    work_name = match.group(2).strip()
                    
                    # 연속된 번호인지 확인 (1, 2, 3...)
                    if int(hopyo_num) == expected_num:
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name,
                            'full_text': cell_str
                        })
                        expected_num += 1
        
        print(f"\n[파싱 결과]")
        print(f"총 {len(hopyo_list)}개 호표 발견")
        
        # 목록과 비교
        if list_info:
            print(f"\n[검증 결과]")
            print(f"총괄표: {len(list_info)}개, 실제 파싱: {len(hopyo_list)}개")
            
            if len(list_info) == len(hopyo_list):
                print("✓ 호표 개수 일치")
            else:
                print(f"✗ 호표 개수 불일치")
            
            # 각 호표별 비교
            for i in range(min(len(list_info), len(hopyo_list))):
                list_item = list_info[i]
                parsed_item = hopyo_list[i]
                
                if list_item['num'] == parsed_item['num']:
                    # 작업명 비교
                    if list_item['work'] in parsed_item['work'] or parsed_item['work'] in list_item['work']:
                        print(f"✓ {list_item['num']}호표: {list_item['work']} = {parsed_item['work'][:40]}...")
                    else:
                        print(f"? {list_item['num']}호표: 총괄표({list_item['work']}) ≠ 파싱({parsed_item['work'][:40]}...)")
                else:
                    print(f"✗ 호표 번호 불일치: 총괄표({list_item['num']}) ≠ 파싱({parsed_item['num']})")
        
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
                # 산출근거 열 확인
                cell = df.iloc[i, sangcul_col]
                if pd.notna(cell):
                    content = str(cell).strip()
                    if content and not content.isspace():
                        # 중분류 패턴 확인
                        sub_match = re.match(r'^(\d+)\.\s*(.+)', content)
                        if sub_match:
                            sub_num = sub_match.group(1)
                            # 대분류가 아닌 경우만 추가
                            if int(sub_num) != int(hopyo['num']):
                                sangcul_content.append({
                                    'row': i,
                                    'type': 'sub_item',
                                    'num': sub_num,
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
        output_file = 'est_sangcul_parsed_v2.json'
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
    parse_est_sangcul()