"""
목록표 기반 통합 단가산출 파서
목록표의 정보와 실제 산출근거 데이터를 정확히 매칭
"""
import pandas as pd
import json
import re
import sys
import io
import os

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

class IntegratedSangculParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.list_info = []
        self.sangcul_data = []
        
    def find_list_sheet(self):
        """단가산출 목록 시트 찾기"""
        try:
            if self.file_path.endswith('.xls'):
                xls = pd.ExcelFile(self.file_path, engine='xlrd')
            else:
                xls = pd.ExcelFile(self.file_path)
            
            list_sheets = []
            sangcul_sheets = []
            
            for sheet in xls.sheet_names:
                if '단가산출' in sheet:
                    if '목록' in sheet or '총괄' in sheet:
                        list_sheets.append(sheet)
                    elif '산근' in sheet or '단가산출서' in sheet or sheet == '단가산출':
                        sangcul_sheets.append(sheet)
                elif '중기단가' in sheet:
                    if '목록' in sheet:
                        list_sheets.append(sheet)
                    elif '산출' in sheet:
                        sangcul_sheets.append(sheet)
                elif '일위대가_산근' in sheet:  # ebs 케이스
                    sangcul_sheets.append(sheet)
                elif '일위대가목록' in sheet:  # ebs 케이스
                    list_sheets.append(sheet)
            
            return list_sheets, sangcul_sheets
            
        except Exception as e:
            print(f"시트 탐색 오류: {str(e)}")
            return [], []
    
    def read_list_info(self, list_sheet):
        """목록 정보 읽기"""
        try:
            if self.file_path.endswith('.xls'):
                df = pd.read_excel(self.file_path, sheet_name=list_sheet, header=None, engine='xlrd')
            else:
                df = pd.read_excel(self.file_path, sheet_name=list_sheet, header=None)
            
            items = []
            
            # 파일별 목록 형식 처리
            if 'test1' in self.file_name:
                # 제N호표 형식
                for i in range(len(df)):
                    cell = df.iloc[i, 0] if 0 < len(df.columns) else None
                    if pd.notna(cell):
                        match = re.match(r'제(\d+)호표', str(cell))
                        if match:
                            num = match.group(1)
                            work = df.iloc[i, 1] if 1 < len(df.columns) else ''
                            spec = df.iloc[i, 2] if 2 < len(df.columns) else ''
                            unit = df.iloc[i, 3] if 3 < len(df.columns) else ''
                            
                            items.append({
                                'num': num,
                                'work': str(work).strip() if pd.notna(work) else '',
                                'spec': str(spec).strip() if pd.notna(spec) else '',
                                'unit': str(unit).strip() if pd.notna(unit) else '',
                                'row': i
                            })
            
            elif 'sgs' in self.file_name:
                # 공종 형식 (호표 없음)
                header_found = False
                for i in range(len(df)):
                    cell = df.iloc[i, 0] if 0 < len(df.columns) else None
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        if '공' in cell_str and '종' in cell_str:
                            header_found = True
                            continue
                        
                        if header_found and cell_str and not cell_str.startswith('단가'):
                            spec = df.iloc[i, 1] if 1 < len(df.columns) else ''
                            unit = df.iloc[i, 3] if 3 < len(df.columns) else ''
                            
                            items.append({
                                'num': str(len(items) + 1),
                                'work': cell_str,
                                'spec': str(spec).strip() if pd.notna(spec) else '',
                                'unit': str(unit).strip() if pd.notna(unit) else '',
                                'row': i
                            })
            
            elif 'est' in self.file_name:
                # 숫자만 있는 형식
                for i in range(len(df)):
                    for j in range(len(df.columns)):
                        cell = df.iloc[i, j]
                        if pd.notna(cell) and str(cell).strip().isdigit():
                            num = int(str(cell))
                            if 1 <= num <= 100:
                                work = None
                                for k in range(j+1, len(df.columns)):
                                    next_cell = df.iloc[i, k]
                                    if pd.notna(next_cell) and not str(next_cell).isdigit():
                                        work = str(next_cell).strip()
                                        break
                                
                                if work:
                                    items.append({
                                        'num': str(num),
                                        'work': work,
                                        'spec': '',
                                        'unit': '',
                                        'row': i
                                    })
                                    break
            
            elif 'ebs' in self.file_name:
                # #N 형식
                for i in range(len(df)):
                    for j in range(len(df.columns)):
                        cell = df.iloc[i, j]
                        if pd.notna(cell):
                            cell_str = str(cell).strip()
                            match = re.match(r'^#(\d+)\s+(.+)', cell_str)
                            if match:
                                num = match.group(1)
                                work = match.group(2)
                                spec = df.iloc[i, j+1] if j+1 < len(df.columns) else ''
                                unit = df.iloc[i, j+2] if j+2 < len(df.columns) else ''
                                
                                items.append({
                                    'num': num,
                                    'work': work,
                                    'spec': str(spec).strip() if pd.notna(spec) else '',
                                    'unit': str(unit).strip() if pd.notna(unit) else '',
                                    'row': i
                                })
                                break
            
            return items
            
        except Exception as e:
            print(f"목록 읽기 오류: {str(e)}")
            return []
    
    def parse_sangcul_data(self, sangcul_sheet):
        """산출근거 데이터 파싱"""
        try:
            if self.file_path.endswith('.xls'):
                df = pd.read_excel(self.file_path, sheet_name=sangcul_sheet, header=None, engine='xlrd')
            else:
                df = pd.read_excel(self.file_path, sheet_name=sangcul_sheet, header=None)
            
            # 헤더 찾기
            header_row = None
            sangcul_col = 0
            bigo_col = None
            
            for i in range(min(20, len(df))):
                for j in range(len(df.columns)):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        if '산출' in cell_str and ('근거' in cell_str or '내역' in cell_str):
                            header_row = i
                            sangcul_col = j
                        elif '비고' in cell_str or '비 고' in cell_str:
                            bigo_col = j
                if header_row is not None:
                    break
            
            # 파일별 패턴 정의
            patterns = []
            if 'test1' in self.file_name or 'est' in self.file_name:
                patterns.append((r'^(\d+)\.\s*(.+)', 'dot_style'))
                sangcul_col = 1 if 'est' in self.file_name else 0
            elif 'sgs' in self.file_name:
                patterns.append((r'산근\s*(\d+)\s*호표\s*[：:](.+)', 'sanggun_style'))
            elif 'ebs' in self.file_name:
                patterns.append((r'^#(\d+)\s+(.+)', 'hash_style'))
            elif '건축구조' in self.file_name:
                patterns.append((r'\(\s*산근\s*(\d+)\s*\)', 'paren_style'))
            
            # 호표 찾기
            hopyo_list = []
            for i in range(header_row + 1 if header_row else 0, len(df)):
                for j in range(len(df.columns)):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        
                        for pattern, style in patterns:
                            match = re.search(pattern, cell_str)
                            if match:
                                hopyo_num = match.group(1)
                                
                                # 작업명 추출
                                if style == 'paren_style':
                                    work_name = cell_str[:match.start()].strip()
                                elif len(match.groups()) > 1:
                                    work_name = match.group(2).strip()
                                else:
                                    work_name = ''
                                
                                hopyo_list.append({
                                    'row': i,
                                    'num': hopyo_num,
                                    'work': work_name,
                                    'style': style,
                                    'full_text': cell_str,
                                    'sangcul_col': sangcul_col,
                                    'bigo_col': bigo_col
                                })
                                break
                        
                        if hopyo_list and hopyo_list[-1]['row'] == i:
                            break
            
            return hopyo_list
            
        except Exception as e:
            print(f"산출근거 파싱 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return []
    
    def extract_hopyo_content(self, df, hopyo, next_hopyo_row):
        """호표 범위의 내용 추출"""
        content = []
        sangcul_col = hopyo['sangcul_col']
        bigo_col = hopyo['bigo_col']
        
        for i in range(hopyo['row'] + 1, next_hopyo_row):
            # 산출근거 내용
            cell = df.iloc[i, sangcul_col]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                if cell_str and not all(c in ['-', '=', ' '] for c in cell_str):
                    item = {
                        'row': i,
                        'content': cell_str
                    }
                    
                    # 비고 추가
                    if bigo_col is not None and bigo_col < len(df.columns):
                        bigo = df.iloc[i, bigo_col]
                        if pd.notna(bigo):
                            bigo_str = str(bigo).strip()
                            if bigo_str and bigo_str != '0':
                                item['bigo'] = bigo_str
                    
                    content.append(item)
        
        return content
    
    def process(self):
        """통합 처리"""
        # 시트 찾기
        list_sheets, sangcul_sheets = self.find_list_sheet()
        
        if not sangcul_sheets:
            print("산출근거 시트를 찾을 수 없습니다.")
            return None
        
        # 목록 정보 읽기
        if list_sheets:
            print(f"목록 시트: {list_sheets[0]}")
            self.list_info = self.read_list_info(list_sheets[0])
            print(f"목록에서 {len(self.list_info)}개 항목 발견")
        
        # 산출근거 파싱
        sangcul_sheet = sangcul_sheets[0]
        print(f"산출근거 시트: {sangcul_sheet}")
        hopyo_data = self.parse_sangcul_data(sangcul_sheet)
        print(f"산출근거에서 {len(hopyo_data)}개 호표 발견")
        
        # 전체 데이터 다시 읽기 (내용 추출용)
        if self.file_path.endswith('.xls'):
            df = pd.read_excel(self.file_path, sheet_name=sangcul_sheet, header=None, engine='xlrd')
        else:
            df = pd.read_excel(self.file_path, sheet_name=sangcul_sheet, header=None)
        
        # 결과 생성
        result = {
            'file': self.file_name,
            'list_sheet': list_sheets[0] if list_sheets else None,
            'sangcul_sheet': sangcul_sheet,
            'list_count': len(self.list_info),
            'parsed_count': len(hopyo_data),
            'data': []
        }
        
        # 목록과 산출근거 매칭
        for idx, hopyo in enumerate(hopyo_data):
            # 다음 호표까지의 범위
            next_row = hopyo_data[idx + 1]['row'] if idx + 1 < len(hopyo_data) else len(df)
            
            # 내용 추출
            content = self.extract_hopyo_content(df, hopyo, next_row)
            
            # 목록에서 매칭되는 항목 찾기
            matched_list = None
            for list_item in self.list_info:
                if list_item['num'] == hopyo['num']:
                    matched_list = list_item
                    break
                # 작업명으로도 매칭 시도
                elif (list_item['work'] in hopyo['work'] or 
                      hopyo['work'] in list_item['work']):
                    matched_list = list_item
                    break
            
            # 데이터 구성
            data_item = {
                'hopyo_num': hopyo['num'],
                'work_name': hopyo['work'],
                'from_list': {
                    'work': matched_list['work'] if matched_list else None,
                    'spec': matched_list['spec'] if matched_list else None,
                    'unit': matched_list['unit'] if matched_list else None,
                } if matched_list else None,
                'start_row': hopyo['row'],
                'end_row': next_row,
                'content_count': len(content),
                'content': content
            }
            
            result['data'].append(data_item)
        
        # 검증
        if self.list_info:
            matched_count = sum(1 for d in result['data'] if d['from_list'] is not None)
            result['validation'] = {
                'matched': matched_count,
                'unmatched': len(hopyo_data) - matched_count,
                'match_rate': f"{matched_count/len(hopyo_data)*100:.1f}%" if hopyo_data else "0%"
            }
        
        return result

def main():
    """모든 파일 처리"""
    files = [
        'test1.xlsx',
        'est.xlsx',
        'sgs.xls',
        'ebs.xls',
        '건축구조내역.xlsx'
    ]
    
    all_results = {}
    
    for file_path in files:
        if os.path.exists(file_path):
            print(f"\n{'='*80}")
            print(f"처리 중: {file_path}")
            print('='*80)
            
            parser = IntegratedSangculParser(file_path)
            result = parser.process()
            
            if result:
                all_results[file_path] = result
                
                # 요약 출력
                print(f"\n[요약]")
                print(f"목록: {result['list_count']}개, 파싱: {result['parsed_count']}개")
                if 'validation' in result:
                    print(f"매칭: {result['validation']['matched']}개 ({result['validation']['match_rate']})")
                
                # 상위 3개 데이터 샘플
                print(f"\n[데이터 샘플]")
                for i, data in enumerate(result['data'][:3]):
                    print(f"\n호표{data['hopyo_num']}: {data['work_name']}")
                    if data['from_list']:
                        print(f"  목록: {data['from_list']['work']} / {data['from_list']['spec']} / {data['from_list']['unit']}")
                    print(f"  내용: {data['content_count']}개 항목")
        else:
            print(f"\n파일 없음: {file_path}")
    
    # 전체 결과 저장
    output_file = 'integrated_sangcul_result.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    
    print(f"\n\n전체 결과 저장: {output_file}")

if __name__ == "__main__":
    main()