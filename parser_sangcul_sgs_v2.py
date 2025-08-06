"""
sgs.xls 단가산출 시트 파서 (목록 검증 기능 포함)
스타일: sanggun_style (산근 1 호표：작업명 형식)
"""
import pandas as pd
import json
import re
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def read_sangcul_list(file_path):
    """단가산출 목록 읽기"""
    try:
        # 목록 시트 이름 찾기
        xls = pd.ExcelFile(file_path, engine='xlrd')
        list_sheet = None
        for sheet in xls.sheet_names:
            if '단가산출' in sheet and '목록' in sheet:
                list_sheet = sheet
                break
        
        if not list_sheet:
            return None
        
        # 목록 읽기
        df = pd.read_excel(file_path, sheet_name=list_sheet, header=None, engine='xlrd')
        
        # 공종 정보 추출 (호표가 아닌 작업명으로 나열)
        work_info = []
        header_found = False
        
        for i in range(len(df)):
            cell = df.iloc[i, 0] if 0 < len(df.columns) else None
            if pd.notna(cell):
                cell_str = str(cell).strip()
                
                # 헤더 찾기
                if '공' in cell_str and '종' in cell_str:
                    header_found = True
                    continue
                
                if header_found and cell_str and not cell_str.startswith('단가'):
                    # 작업명
                    spec = df.iloc[i, 1] if 1 < len(df.columns) else ''
                    
                    work_info.append({
                        'work': cell_str,
                        'spec': str(spec).strip() if pd.notna(spec) else ''
                    })
        
        print(f"\n[단가산출 목록 정보]")
        print(f"총 {len(work_info)}개 작업 발견")
        for i, info in enumerate(work_info[:10]):  # 처음 10개만 출력
            print(f"  {i+1}. {info['work']} ({info['spec']})")
        if len(work_info) > 10:
            print(f"  ... 외 {len(work_info)-10}개")
        
        return work_info
        
    except Exception as e:
        print(f"목록 읽기 오류: {str(e)}")
        return None

def parse_sgs_sangcul():
    """sgs.xls의 단가산출 시트 파싱"""
    
    file_path = 'sgs.xls'
    sheet_name = '단가산출'
    
    try:
        # 먼저 목록 정보 읽기
        list_info = read_sangcul_list(file_path)
        
        # 엑셀 파일 읽기 (xls는 xlrd 엔진 사용)
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
        print(f"\n시트 '{sheet_name}' 로드 완료. 크기: {df.shape}")
        
        # 산출근거 열은 0열
        sangcul_col = 0
        
        # 호표 패턴: 산근 N 호표
        hopyo_pattern = r'산근\s*(\d+)\s*호표\s*[：:](.+)'
        
        # 모든 호표 찾기
        hopyo_list = []
        
        for i in range(len(df)):
            cell = df.iloc[i, sangcul_col]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.search(hopyo_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    work_name = match.group(2).strip()
                    
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
            print(f"목록: {len(list_info)}개 작업, 실제 파싱: {len(hopyo_list)}개 호표")
            
            # SGS는 목록이 작업명으로 나열되고, 실제는 호표로 구분되므로
            # 정확한 1:1 매칭은 어려움. 개수 비교만 수행
            if len(list_info) == len(hopyo_list):
                print("✓ 항목 개수 일치")
            else:
                print(f"! 항목 개수 차이 (목록과 호표 구조가 다름)")
            
            # 처음 몇 개만 비교
            print("\n[상위 5개 항목 비교]")
            for i in range(min(5, len(hopyo_list))):
                parsed = hopyo_list[i]
                print(f"  호표{parsed['num']}: {parsed['work']}")
                
                # 목록에서 유사한 항목 찾기
                for list_item in list_info:
                    if parsed['work'] in list_item['work'] or list_item['work'] in parsed['work']:
                        print(f"    → 목록에서 발견: {list_item['work']}")
                        break
        
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
            
            # 중분류 패턴
            sub_pattern = r'^(\d+)\.\s*(.+)'
            
            for i in range(start_row, end_row):
                cell = df.iloc[i, sangcul_col]
                if pd.notna(cell):
                    content = str(cell).strip()
                    if content and not content.isspace():
                        # 호표 행은 제외
                        if i == start_row:
                            continue
                        
                        # 중분류 패턴 확인
                        sub_match = re.match(sub_pattern, content)
                        if sub_match:
                            sub_num = sub_match.group(1)
                            sub_name = sub_match.group(2)
                            sangcul_content.append({
                                'row': i,
                                'type': 'sub_item',
                                'num': sub_num,
                                'name': sub_name,
                                'content': content
                            })
                        else:
                            # 일반 내용 (공백이 많은 구분선 제외)
                            if not all(c in ['-', ' ', '='] for c in content):
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
        output_file = 'sgs_sangcul_parsed_v2.json'
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
    parse_sgs_sangcul()