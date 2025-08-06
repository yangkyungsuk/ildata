"""
ebs.xls 일위대가_산근 시트 파서
스타일: hash_style (#1 작업명 형식)
특징: 비고 열 포함
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def parse_ebs_sangcul():
    """ebs.xls의 일위대가_산근 시트 파싱"""
    
    file_path = 'ebs.xls'
    sheet_name = '일위대가_산근'
    
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
        print(f"시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
        # 헤더 찾기
        header_row = None
        sangcul_col = 0
        bigo_col = None
        
        for i in range(min(10, len(df))):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    if '산' in cell_str and '출' in cell_str and '근' in cell_str:
                        header_row = i
                        sangcul_col = j
                    elif '비' in cell_str and '고' in cell_str:
                        bigo_col = j
            if header_row is not None:
                break
        
        print(f"헤더 행: {header_row}, 산출근거 열: {sangcul_col}, 비고 열: {bigo_col}")
        
        # 호표 패턴: #숫자 작업명
        hopyo_pattern = r'^#(\d+)\s+(.+)'
        
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
                    
                    # 파이프(|)로 구분된 경우 처리
                    if '|' in work_name:
                        parts = work_name.split('|')
                        work_name = parts[0].strip()
                    
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
        
        # 중분류 패턴
        sub_pattern = r'^(\d+)\.\s*(.+)'
        
        for idx, hopyo in enumerate(hopyo_list):
            hopyo_key = f"호표{hopyo['num']}"
            
            # 다음 호표까지의 범위
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else min(start_row + 100, len(df))
            
            # 산출근거 내용 수집
            sangcul_content = []
            
            for i in range(start_row + 1, end_row):
                # 산출근거 내용
                cell = df.iloc[i, sangcul_col]
                bigo = df.iloc[i, bigo_col] if bigo_col and bigo_col < len(df.columns) else None
                
                if pd.notna(cell):
                    content = str(cell).strip()
                    if content and not content.isspace():
                        # 데이터 구조
                        data_item = {
                            'row': i,
                            'content': content
                        }
                        
                        # 비고 추가
                        if pd.notna(bigo):
                            bigo_str = str(bigo).strip()
                            if bigo_str and bigo_str != '0':
                                data_item['bigo'] = bigo_str
                        
                        # 중분류 패턴 확인
                        sub_match = re.match(sub_pattern, content)
                        if sub_match:
                            data_item['type'] = 'sub_item'
                            data_item['num'] = sub_match.group(1)
                            data_item['name'] = sub_match.group(2)
                        else:
                            data_item['type'] = 'content'
                        
                        sangcul_content.append(data_item)
            
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
        output_file = 'ebs_sangcul_parsed.json'
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
    parse_ebs_sangcul()