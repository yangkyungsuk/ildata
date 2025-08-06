"""
test1.xlsx 단가산출_산근 시트 파서 v3 (호표 사이 모든 데이터 수집)
스타일: dot_style (1.작업명 형식) + 세부 내역 완전 수집
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
        xls = pd.ExcelFile(file_path)
        list_sheet = None
        for sheet in xls.sheet_names:
            if '단가산출' in sheet and '목록' in sheet:
                list_sheet = sheet
                break
        
        if not list_sheet:
            return None
        
        df = pd.read_excel(file_path, sheet_name=list_sheet, header=None)
        
        hopyo_info = []
        for i in range(len(df)):
            cell = df.iloc[i, 0] if 0 < len(df.columns) else None
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(r'제(\d+)호표', cell_str)
                if match:
                    hopyo_num = match.group(1)
                    work_name = df.iloc[i, 1] if 1 < len(df.columns) else ''
                    spec = df.iloc[i, 2] if 2 < len(df.columns) else ''
                    
                    hopyo_info.append({
                        'num': hopyo_num,
                        'work': str(work_name).strip() if pd.notna(work_name) else '',
                        'spec': str(spec).strip() if pd.notna(spec) else ''
                    })
        
        print(f"\n[단가산출목록표 정보]")
        print(f"총 {len(hopyo_info)}개 호표 발견")
        
        return hopyo_info
        
    except Exception as e:
        print(f"목록 읽기 오류: {str(e)}")
        return None

def extract_detailed_content(df, start_row, end_row):
    """호표 범위 내 모든 세부 데이터 추출"""
    detailed_data = []
    
    for row in range(start_row, end_row):
        row_data = {'row': row, 'columns': {}}
        has_content = False
        
        # 모든 컬럼 검사
        for col in range(min(15, len(df.columns))):  # 처음 15컬럼까지
            cell = df.iloc[row, col]
            if pd.notna(cell):
                content = str(cell).strip()
                if content and content not in ['', '0', '0.0']:
                    row_data['columns'][f'col_{col}'] = content
                    has_content = True
        
        if has_content:
            # 데이터 타입 분류
            main_content = row_data['columns'].get('col_0', '')
            
            # 호표 헤더 (1.작업명 형식)
            if re.match(r'^\d+\.\s*[가-힣]', main_content):
                row_data['type'] = 'hopyo_header'
                row_data['description'] = main_content
            
            # 하위 작업 항목 (1., 2., 3. 등)
            elif re.match(r'^\s*\d+\.\s*[가-힣]', main_content):
                row_data['type'] = 'sub_work'
                row_data['description'] = main_content.strip()
            
            # 계산 변수 (Q =, n =, t = 등)
            elif any('=' in str(v) for v in row_data['columns'].values()):
                row_data['type'] = 'calculation'
                # = 기호가 있는 컬럼 찾기
                for col_key, value in row_data['columns'].items():
                    if '=' in str(value):
                        parts = str(value).split('=')
                        if len(parts) == 2:
                            row_data['variable'] = parts[0].strip()
                            row_data['value'] = parts[1].strip()
                            break
            
            # 비용 항목 (재료비, 노무비, 경비, 소계)
            elif any(keyword in main_content for keyword in ['재료비', '노무비', '경비', '소계', '합계']):
                row_data['type'] = 'cost_item'
                row_data['cost_type'] = main_content
                # 금액이 있는 컬럼 찾기
                for col_key, value in row_data['columns'].items():
                    if col_key != 'col_0' and re.match(r'[\d,\.]+', str(value)):
                        row_data['amount'] = str(value)
                        break
            
            # 장비/인력 정보
            elif any(keyword in main_content for keyword in ['브레이카', '굴삭기', '백호', '인부', '운전수', '기사']):
                row_data['type'] = 'equipment_labor'
                row_data['description'] = main_content
            
            # 기타 설명
            else:
                row_data['type'] = 'description'
                row_data['description'] = main_content
            
            detailed_data.append(row_data)
    
    return detailed_data

def parse_test1_sangcul():
    """test1.xlsx의 단가산출_산근 시트 파싱 (세부 데이터 완전 수집)"""
    
    file_path = 'test1.xlsx'
    sheet_name = '단가산출_산근'
    
    try:
        # 목록 정보 읽기
        list_info = read_sangcul_list(file_path)
        
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        print(f"\n시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
        # 헤더 찾기
        header_row = None
        sangcul_col = 0
        
        for i in range(min(10, len(df))):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell) and '산출근거' in str(cell):
                    header_row = i
                    sangcul_col = j
                    break
            if header_row is not None:
                break
        
        # 호표 패턴: 숫자.작업명 형식
        hopyo_pattern = r'^(\d+)\.\s*(.+)'
        
        # 호표 찾기
        hopyo_list = []
        for i in range(header_row + 1 if header_row else 0, len(df)):
            cell = df.iloc[i, sangcul_col]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(hopyo_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    work_name = match.group(2).strip()
                    
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
            print("✓ 호표 개수 일치" if len(list_info) == len(hopyo_list) else "✗ 호표 개수 불일치")
        
        # 각 호표의 세부 데이터 추출
        result = {
            'file': file_path,
            'sheet': sheet_name,
            'hopyo_count': len(hopyo_list),
            'validation': {
                'list_count': len(list_info) if list_info else 0,
                'parsed_count': len(hopyo_list),
                'match': len(list_info) == len(hopyo_list) if list_info else False
            },
            'hopyo_data': {}
        }
        
        for idx, hopyo in enumerate(hopyo_list):
            hopyo_key = f"호표{hopyo['num']}"
            
            # 호표 범위
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(df)
            
            # 세부 데이터 추출
            detailed_content = extract_detailed_content(df, start_row, end_row)
            
            # 데이터 타입별 분류
            sub_works = [item for item in detailed_content if item.get('type') == 'sub_work']
            calculations = [item for item in detailed_content if item.get('type') == 'calculation']
            cost_items = [item for item in detailed_content if item.get('type') == 'cost_item']
            equipment_labor = [item for item in detailed_content if item.get('type') == 'equipment_labor']
            
            result['hopyo_data'][hopyo_key] = {
                '호표번호': hopyo['num'],
                '작업명': hopyo['work'],
                '시작행': start_row,
                '종료행': end_row,
                '총_데이터_행수': len(detailed_content),
                '하위작업': sub_works,
                '계산변수': calculations,
                '비용항목': cost_items,
                '장비인력': equipment_labor,
                '전체_세부데이터': detailed_content
            }
            
            print(f"\n호표{hopyo['num']}: {hopyo['work']}")
            print(f"  - 총 {len(detailed_content)}행 데이터")
            print(f"  - 하위작업: {len(sub_works)}개")
            print(f"  - 계산변수: {len(calculations)}개")
            print(f"  - 비용항목: {len(cost_items)}개")
            print(f"  - 장비인력: {len(equipment_labor)}개")
        
        # JSON 파일로 저장
        output_file = 'test1_sangcul_parsed_v3.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n파싱 완료. 결과 저장: {output_file}")
        print(f"호표 사이의 모든 세부 데이터가 수집되었습니다.")
        
        return result
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    parse_test1_sangcul()