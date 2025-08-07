"""
통합 일위대가 파서
모든 엑셀 파일 타입에 대응하는 개선된 파서
"""
import pandas as pd
import json
import os
import re
from typing import Dict, List, Tuple, Optional

class UnifiedIlwidaeParser:
    """통합 일위대가 파서 클래스"""
    
    def __init__(self):
        # result 폴더 생성
        os.makedirs('result', exist_ok=True)
        
        # 파일별 설정
        self.file_configs = {
            'test1.xlsx': {
                'sheet': '일위대가_산근',
                'pattern': r'제\s*(\d+)\s*호표',
                'header_row': 3,
                'col_mapping': {'품명': 1, '규격': 2, '단위': 3, '수량': 4}
            },
            'est.xlsx': {
                'sheet': '일위대가',
                'pattern': r'제\s*(\d+)\s*호표',
                'header_row': 3,
                'col_mapping': None  # 자동 감지
            },
            'sgs.xls': {
                'sheet': '일위대가',
                'pattern': r'제\s*(\d+)\s*호표',
                'header_row': 3,
                'col_mapping': None  # 자동 감지
            },
            '건축구조내역.xlsx': {
                'sheet': '일위대가',
                'pattern': r'[A-Z]-(\d+)',  # M-101 패턴
                'header_row': 2,
                'col_mapping': None  # 자동 감지
            },
            '건축구조내역2.xlsx': {
                'sheet': '일위대가', 
                'pattern': r'[A-Z]-(\d+)',  # M-101 패턴
                'header_row': 2,
                'col_mapping': None  # 자동 감지
            }
        }
    
    def detect_columns(self, df: pd.DataFrame, start_row: int = 0) -> Dict:
        """컬럼 위치 자동 감지"""
        column_info = {
            '품명': None,
            '규격': None,
            '단위': None,
            '수량': None
        }
        
        # 헤더 행 찾기
        for row_idx in range(start_row, min(start_row + 10, len(df))):
            row_data = []
            for col_idx in range(min(15, len(df.columns))):
                cell = df.iloc[row_idx, col_idx]
                if pd.notna(cell):
                    row_data.append(str(cell).strip())
            
            # 헤더 키워드 찾기
            row_text = ' '.join(row_data).lower()
            if '품' in row_text and ('명' in row_text or '종' in row_text):
                # 헤더 행 발견
                for col_idx, value in enumerate(row_data):
                    value_lower = value.lower()
                    if '품' in value_lower or '공종' in value_lower:
                        column_info['품명'] = col_idx
                    elif '규' in value_lower or '격' in value_lower:
                        column_info['규격'] = col_idx
                    elif '단' in value_lower and '위' in value_lower:
                        column_info['단위'] = col_idx
                    elif '수' in value_lower and '량' in value_lower:
                        column_info['수량'] = col_idx
                break
        
        # 헤더를 못 찾았으면 데이터 패턴으로 추정
        if column_info['품명'] is None:
            for col_idx in range(min(10, len(df.columns))):
                col_values = []
                for row_idx in range(start_row + 5, min(start_row + 15, len(df))):
                    cell = df.iloc[row_idx, col_idx]
                    if pd.notna(cell):
                        col_values.append(str(cell).strip())
                
                if col_values:
                    # 한글이 많으면 품명
                    if column_info['품명'] is None and any(re.match(r'^[가-힣]+', v) for v in col_values):
                        column_info['품명'] = col_idx
                    # 단위 키워드
                    elif any(v in ['인', '%', 'M3', 'M2', 'M', 'KG', 'TON', '개', '대', 'HR', '조', 'm3', 'L', '식'] for v in col_values):
                        column_info['단위'] = col_idx
                    # 작은 숫자는 수량
                    elif column_info['수량'] is None:
                        nums = [float(v) for v in col_values if re.match(r'^\d+\.?\d*$', v)]
                        if nums and all(n < 1000 for n in nums):
                            column_info['수량'] = col_idx
        
        # 규격은 품명과 단위 사이로 추정
        if column_info['품명'] is not None and column_info['단위'] is not None:
            if column_info['규격'] is None:
                column_info['규격'] = column_info['품명'] + 1
        
        return column_info
    
    def find_hopyos(self, df: pd.DataFrame, pattern: str) -> List[Dict]:
        """호표 찾기"""
        hopyo_list = []
        seen_nums = set()
        
        for i in range(len(df)):
            for j in range(min(10, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # 패턴 매칭
                    match = re.search(pattern, cell_str)
                    if match:
                        hopyo_num = match.group(1)
                        
                        # 중복 체크
                        if hopyo_num in seen_nums:
                            continue
                        
                        # 작업명 찾기
                        work_name = ''
                        
                        # 같은 셀에 작업명이 있는 경우 (제1호표：작업명)
                        if '：' in cell_str or ':' in cell_str:
                            parts = re.split('[：:]', cell_str)
                            if len(parts) > 1:
                                work_name = parts[1].strip()
                        
                        # 다음 컬럼에서 작업명 찾기
                        if not work_name:
                            for k in range(j+1, min(j+5, len(df.columns))):
                                next_cell = df.iloc[i, k]
                                if pd.notna(next_cell):
                                    next_str = str(next_cell).strip()
                                    if next_str and not re.match(r'^[\d.,]+$', next_str):
                                        work_name = next_str
                                        break
                        
                        # 작업명이 없으면 다음 행에서 찾기
                        if not work_name and i + 1 < len(df):
                            for k in range(min(5, len(df.columns))):
                                next_cell = df.iloc[i + 1, k]
                                if pd.notna(next_cell):
                                    next_str = str(next_cell).strip()
                                    if next_str and not re.match(r'^[\d.,]+$', next_str) and '합계' not in next_str:
                                        work_name = next_str
                                        break
                        
                        if work_name:
                            seen_nums.add(hopyo_num)
                            hopyo_list.append({
                                'row': i,
                                'num': hopyo_num,
                                'work': work_name
                            })
                        break
        
        return hopyo_list
    
    def extract_sangul_items(self, df: pd.DataFrame, start_row: int, end_row: int, 
                            column_info: Dict) -> List[Dict]:
        """산출근거 항목 추출"""
        items = []
        
        for row_idx in range(start_row + 1, min(end_row, len(df))):
            row_data = {
                'row_number': row_idx,
                '품명': '',
                '규격': '',
                '단위': '',
                '수량': ''
            }
            
            # 각 컬럼에서 데이터 추출
            has_content = False
            
            # 품명
            if column_info['품명'] is not None:
                cell = df.iloc[row_idx, column_info['품명']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    # 합계, 계 등은 제외
                    if value and '합계' not in value and '계' != value:
                        row_data['품명'] = value
                        has_content = True
            
            # 규격
            if column_info['규격'] is not None:
                cell = df.iloc[row_idx, column_info['규격']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data['규격'] = value
            
            # 단위
            if column_info['단위'] is not None:
                cell = df.iloc[row_idx, column_info['단위']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data['단위'] = value
            
            # 수량
            if column_info['수량'] is not None:
                cell = df.iloc[row_idx, column_info['수량']]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data['수량'] = value
            
            if has_content:
                items.append(row_data)
        
        return items
    
    def parse_file(self, file_path: str) -> Optional[Dict]:
        """파일 파싱"""
        if not os.path.exists(file_path):
            print(f"파일 없음: {file_path}")
            return None
        
        config = self.file_configs.get(file_path)
        if not config:
            print(f"설정 없음: {file_path}")
            return None
        
        print(f"\n{'='*70}")
        print(f"파싱 중: {file_path}")
        print('='*70)
        
        try:
            # 파일 읽기
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, sheet_name=config['sheet'], header=None, engine='xlrd')
            else:
                df = pd.read_excel(file_path, sheet_name=config['sheet'], header=None)
            
            print(f"시트 크기: {df.shape}")
            
            # 컬럼 정보 가져오기
            if config['col_mapping']:
                column_info = config['col_mapping']
            else:
                column_info = self.detect_columns(df, config['header_row'])
            
            print(f"컬럼 매핑: {column_info}")
            
            # 호표 찾기
            hopyo_list = self.find_hopyos(df, config['pattern'])
            print(f"호표 발견: {len(hopyo_list)}개")
            
            # 결과 구성
            result = {
                'file': file_path,
                'sheet': config['sheet'],
                'total_ilwidae_count': len(hopyo_list),
                'ilwidae_data': []
            }
            
            # 각 호표 처리
            for idx, hopyo in enumerate(hopyo_list):
                start_row = hopyo['row']
                end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(df)
                
                # 산출근거 추출
                sangul_items = self.extract_sangul_items(df, start_row, end_row, column_info)
                
                ilwidae_item = {
                    'ilwidae_no': hopyo['num'],
                    'ilwidae_title': {
                        '품명': hopyo['work'],
                        '규격': '',
                        '단위': '',
                        '수량': ''
                    },
                    'position': {
                        'start_row': start_row,
                        'end_row': end_row,
                        'total_rows': end_row - start_row
                    },
                    '산출근거': sangul_items
                }
                
                result['ilwidae_data'].append(ilwidae_item)
                
                if idx < 3:  # 처음 3개만 출력
                    print(f"  호표 {hopyo['num']}: {hopyo['work'][:30]} ({len(sangul_items)}개 항목)")
            
            # 저장
            output_file = f"result/{file_path.replace('.', '_')}_ilwidae_unified.json"
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            
            print(f"\n✅ 저장 완료: {output_file}")
            
            # 통계
            total_sangul = sum(len(item['산출근거']) for item in result['ilwidae_data'])
            print(f"   총 산출근거 항목: {total_sangul}개")
            
            return result
            
        except Exception as e:
            print(f"오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

def main():
    """메인 실행"""
    parser = UnifiedIlwidaeParser()
    
    files = [
        'test1.xlsx',
        'est.xlsx',
        'sgs.xls',
        '건축구조내역.xlsx',
        '건축구조내역2.xlsx'
    ]
    
    results = []
    for file_path in files:
        result = parser.parse_file(file_path)
        if result:
            results.append(result)
    
    # 전체 요약
    print("\n" + "="*70)
    print("전체 파싱 요약")
    print("="*70)
    
    for result in results:
        print(f"\n{result['file']}:")
        print(f"  - 일위대가: {result['total_ilwidae_count']}개")
        total_sangul = sum(len(item['산출근거']) for item in result['ilwidae_data'])
        print(f"  - 산출근거: {total_sangul}개")
        
        # 필드 채움 통계
        filled_counts = {'품명': 0, '규격': 0, '단위': 0, '수량': 0}
        total_items = 0
        
        for item in result['ilwidae_data']:
            for sangul in item['산출근거']:
                total_items += 1
                for field in filled_counts.keys():
                    if sangul.get(field, ''):
                        filled_counts[field] += 1
        
        if total_items > 0:
            print(f"  - 필드 채움율:")
            for field, count in filled_counts.items():
                percentage = count / total_items * 100
                print(f"    {field}: {percentage:.1f}%")

if __name__ == "__main__":
    main()