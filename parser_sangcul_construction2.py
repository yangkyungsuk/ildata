"""
건축구조내역2.xlsx 단가산출서 시트 파서
스타일: 분석 필요
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def parse_construction2_sangcul():
    """건축구조내역2.xlsx의 단가산출서 시트 파싱"""
    
    file_path = '건축구조내역2.xlsx'
    sheet_name = '단가산출서'
    
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        print(f"시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
        # 헤더 찾기
        header_row = None
        sangcul_col = None
        
        for i in range(min(20, len(df))):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    if '산출' in cell_str and ('근거' in cell_str or '내역' in cell_str):
                        header_row = i
                        sangcul_col = j
                        print(f"헤더 행 발견: {header_row}, 산출 열: {sangcul_col}")
                        break
            if header_row is not None:
                break
        
        # 패턴 정의
        patterns = [
            (r'^(\d+)\.\s*(.+)', 'dot_style'),  # 1.작업명
            (r'산근\s*(\d+)\s*호표\s*[：:](.+)', 'sanggun_style'),  # 산근 호표
            (r'^(\d+)\s*(.+)', 'number_style'),  # 숫자만 (건축구조 스타일)
        ]
        
        # 호표 찾기
        hopyo_list = []
        expected_num = 1
        
        start_row = header_row + 1 if header_row else 0
        for i in range(start_row, len(df)):
            # 각 열을 확인 (처음 5열만)
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # 숫자만 있는 경우 체크
                    if cell_str.isdigit() and int(cell_str) == expected_num:
                        # 다음 열에서 작업명 찾기
                        work_name = ''
                        for k in range(j+1, min(j+5, len(df.columns))):
                            next_cell = df.iloc[i, k]
                            if pd.notna(next_cell):
                                work_name = str(next_cell).strip()
                                if work_name:
                                    break
                        
                        if work_name:
                            hopyo_list.append({
                                'row': i,
                                'num': str(expected_num),
                                'work': work_name,
                                'style': 'number_style'
                            })
                            expected_num += 1
                            break
                    
                    # 다른 패턴 확인
                    for pattern, style in patterns:
                        match = re.match(pattern, cell_str)
                        if match and len(match.groups()) >= 2:
                            num = match.group(1)
                            work = match.group(2)
                            if int(num) == expected_num:
                                hopyo_list.append({
                                    'row': i,
                                    'num': num,
                                    'work': work,
                                    'style': style
                                })
                                expected_num += 1
                                break
        
        print(f"총 {len(hopyo_list)}개 호표 발견")
        
        # 산출근거 열 결정
        if sangcul_col is None:
            sangcul_col = 1  # 기본값
        
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
            
            for i in range(start_row + 1, end_row):
                # 여러 열에서 내용 수집
                row_content = []
                for j in range(len(df.columns)):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        content = str(cell).strip()
                        if content and not content.replace('.', '').replace(' ', '').isdigit():
                            row_content.append(content)
                
                if row_content:
                    # 전체 행 내용을 합침
                    full_content = ' | '.join(row_content)
                    sangcul_content.append({
                        'row': i,
                        'type': 'content',
                        'content': full_content
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
        output_file = 'construction2_sangcul_parsed.json'
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
    parse_construction2_sangcul()