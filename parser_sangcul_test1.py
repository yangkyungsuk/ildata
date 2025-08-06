"""
test1.xlsx 단가산출_산근 시트 파서
스타일: dot_style (1.작업명 형식)
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def parse_test1_sangcul():
    """test1.xlsx의 단가산출_산근 시트 파싱"""
    
    file_path = 'test1.xlsx'
    sheet_name = '단가산출_산근'
    
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        print(f"시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
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
        
        print(f"총 {len(hopyo_list)}개 호표 발견")
        
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
            
            print(f"{hopyo_key}: {hopyo['work']} - {len(sangcul_content)}개 항목")
        
        # JSON 파일로 저장
        output_file = 'test1_sangcul_parsed.json'
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