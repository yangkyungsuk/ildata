"""
건축구조내역.xlsx 중기단가산출서 시트 파서
스타일: 분석 필요
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def parse_construction_sangcul():
    """건축구조내역.xlsx의 중기단가산출서 시트 파싱"""
    
    file_path = '건축구조내역.xlsx'
    sheet_name = '중기단가산출서'
    
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
        
        # 패턴: ( 산근 N ) 형식
        pattern = r'\(\s*산근\s*(\d+)\s*\)'
        
        # 호표 찾기
        hopyo_list = []
        
        start_row = header_row + 1 if header_row else 0
        for i in range(start_row, len(df)):
            # 각 열을 확인
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # 산근 패턴 찾기
                    match = re.search(pattern, cell_str)
                    if match:
                        hopyo_num = match.group(1)
                        
                        # 작업명은 산근 앞부분
                        work_name = cell_str[:match.start()].strip()
                        
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name,
                            'full_text': cell_str,
                            'style': 'sanggun_paren'  # ( 산근 N ) 스타일
                        })
                        break
        
        print(f"총 {len(hopyo_list)}개 호표 발견")
        
        # 산출근거 열 결정
        if sangcul_col is None:
            sangcul_col = 0  # 기본값
        
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
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else min(start_row + 100, len(df))
            
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
        output_file = 'construction_sangcul_parsed.json'
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
    parse_construction_sangcul()