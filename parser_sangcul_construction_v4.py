"""
건축구조내역.xlsx 중기단가산출서 시트 파서 v4 (완전한 순서 보존)
규칙: 호표 사이의 모든 행을 순서대로 빠짐없이 JSON에 저장
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def read_sangcul_list(file_path):
    """목록 읽기 (건축구조내역용)"""
    try:
        # 시트 이름 탐색
        xls = pd.ExcelFile(file_path)
        list_sheet = None
        for sheet in xls.sheet_names:
            if '목록' in sheet or '총괄' in sheet or '중기단가' in sheet:
                if '산출' not in sheet:  # 산출서가 아닌 목록/총괄 시트
                    list_sheet = sheet
                    break
        
        if not list_sheet:
            print("목록 시트를 찾을 수 없습니다.")
            return None
        
        df = pd.read_excel(file_path, sheet_name=list_sheet, header=None)
        
        hopyo_info = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    # (산근 N) 패턴 찾기
                    match = re.search(r'\(\s*산근\s*(\d+)\s*\)', cell_str)
                    if match:
                        num = match.group(1)
                        work = cell_str[:match.start()].strip()
                        
                        if work:
                            hopyo_info.append({
                                'num': num,
                                'work': work,
                                'spec': '',
                                'unit': ''
                            })
                            break
        
        print(f"\n[목록 정보]")
        print(f"총 {len(hopyo_info)}개 호표 발견")
        
        return hopyo_info
        
    except Exception as e:
        print(f"목록 읽기 오류: {str(e)}")
        return None

def extract_all_rows_in_order(df, start_row, end_row):
    """호표 사이의 모든 행을 순서대로 완전히 추출"""
    all_rows = []
    
    for row_idx in range(start_row, end_row):
        # 해당 행의 모든 컬럼 데이터 수집
        row_data = {
            'row_number': row_idx,
            'columns': {},
            'has_content': False
        }
        
        # 모든 컬럼 검사 (최대 20컬럼)
        for col_idx in range(min(20, len(df.columns))):
            cell = df.iloc[row_idx, col_idx]
            if pd.notna(cell):
                content = str(cell).strip()
                if content:  # 빈 문자열이 아니면 저장
                    row_data['columns'][f'col_{col_idx}'] = content
                    row_data['has_content'] = True
        
        # 빈 행이어도 순서 유지를 위해 모두 저장
        all_rows.append(row_data)
    
    return all_rows

def parse_construction_sangcul():
    """건축구조내역.xlsx의 중기단가산출서 시트 파싱 (완전한 순서 보존)"""
    
    file_path = '건축구조내역.xlsx'
    sheet_name = '중기단가산출서'
    
    try:
        # 목록 정보 읽기
        list_info = read_sangcul_list(file_path)
        
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        print(f"\n시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
        # 호표 패턴: 작업명 (산근 N)
        hopyo_pattern = r'(.+)\(\s*산근\s*(\d+)\s*\)'
        
        # 호표 찾기
        hopyo_list = []
        for i in range(len(df)):
            for j in range(min(3, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    match = re.search(hopyo_pattern, cell_str)
                    if match:
                        work_name = match.group(1).strip()
                        hopyo_num = match.group(2)
                        
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name,
                            'full_text': cell_str
                        })
                        break
        
        print(f"\n[파싱 결과]")
        print(f"총 {len(hopyo_list)}개 호표 발견")
        
        # 목록과 비교
        if list_info:
            print(f"\n[검증 결과]")
            print(f"목록표: {len(list_info)}개, 실제 파싱: {len(hopyo_list)}개")
            print("✓ 호표 개수 일치" if len(list_info) == len(hopyo_list) else "✗ 호표 개수 불일치")
        
        # 결과 구성
        result = {
            'file': file_path,
            'sheet': sheet_name,
            'total_hopyo_count': len(hopyo_list),
            'validation': {
                'list_count': len(list_info) if list_info else 0,
                'parsed_count': len(hopyo_list),
                'match': len(list_info) == len(hopyo_list) if list_info else False
            },
            'hopyo_data': {}
        }
        
        # 각 호표의 완전한 데이터 추출
        for idx, hopyo in enumerate(hopyo_list):
            hopyo_key = f"호표{hopyo['num']}"
            
            # 호표 범위 설정
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(df)
            
            # 호표 사이의 모든 행을 순서대로 완전히 추출
            all_rows_data = extract_all_rows_in_order(df, start_row, end_row)
            
            # 내용이 있는 행만 계산 (통계용)
            content_rows = [row for row in all_rows_data if row['has_content']]
            
            result['hopyo_data'][hopyo_key] = {
                '호표번호': hopyo['num'],
                '작업명': hopyo['work'],
                '시작행': start_row,
                '종료행': end_row,
                '총_행수': len(all_rows_data),
                '내용있는_행수': len(content_rows),
                '모든_행_데이터': all_rows_data  # 순서대로 모든 행 저장
            }
            
            print(f"\n호표{hopyo['num']}: {hopyo['work']}")
            print(f"  - 전체 행 범위: {start_row} ~ {end_row-1} ({len(all_rows_data)}개 행)")
            print(f"  - 내용이 있는 행: {len(content_rows)}개")
            
            # 처음 5개 행 샘플 출력
            print(f"  - 처음 5개 행 샘플:")
            for i, row_data in enumerate(all_rows_data[:5]):
                if row_data['has_content']:
                    main_content = list(row_data['columns'].values())[0] if row_data['columns'] else ''
                    print(f"    행{row_data['row_number']}: \"{main_content[:40]}...\"")
                else:
                    print(f"    행{row_data['row_number']}: (빈행)")
        
        # JSON 파일로 저장
        output_file = 'construction_sangcul_parsed_v4.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n파싱 완료. 결과 저장: {output_file}")
        print(f"✅ 호표 사이의 모든 행이 순서대로 완전히 저장되었습니다.")
        
        return result
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    parse_construction_sangcul()