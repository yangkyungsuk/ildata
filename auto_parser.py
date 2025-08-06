import pandas as pd
import json
import warnings
import sys
import io
import re
import os
from typing import Dict, Any, Tuple

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')

class AutoParser:
    """파일 구조를 자동으로 판별하고 적절한 파서를 실행하는 클래스"""
    
    def __init__(self):
        self.file_patterns = {
            'test1_type': {
                'header_row': 1,
                'columns': ['호표', '품명', '규격', '단위', '수량'],
                'hopyo_format': '제N호표',
                'description': 'test1 형식 (호표|품명|규격|단위|수량)'
            },
            'sgs_type': {
                'header_row': 2,
                'columns': ['공종', '규격', '수량', '단위'],
                'hopyo_format': '제 N 호표 : 작업명',
                'description': 'sgs 형식 (공종|규격|수량|단위, 호표 통합)'
            },
            'construction_type': {
                'header_row': 1,
                'columns': ['품명', '규격', '단위', '수량'],
                'hopyo_format': '작업명 (호표 N)',
                'description': '건축구조 형식 (품명|규격|단위|수량, 호표 통합)'
            },
            'est_type': {
                'header_row': 2,
                'columns': ['호표', '품명', '규격', '수량', '단위'],
                'hopyo_format': '제N호표',
                'description': 'est 형식 (호표|품명|규격|수량|단위)'
            }
        }
    
    def detect_file_type(self, file_path: str) -> Tuple[str, Dict]:
        """파일 구조를 자동으로 감지"""
        
        print(f"\n파일 구조 감지 중: {file_path}")
        print("-" * 60)
        
        # 시트 목록 확인
        if file_path.endswith('.xls'):
            xls = pd.ExcelFile(file_path, engine='xlrd')
        else:
            xls = pd.ExcelFile(file_path)
        
        sheet_names = xls.sheet_names
        print(f"시트 목록: {sheet_names}")
        
        # 일위대가 관련 시트 찾기 (목록 제외)
        target_sheet = None
        for sheet in sheet_names:
            if '일위대가' in sheet and '목록' not in sheet and '총괄' not in sheet:
                target_sheet = sheet
                break
        
        # 못 찾으면 산근 시트 찾기
        if not target_sheet:
            for sheet in sheet_names:
                if '산근' in sheet:
                    target_sheet = sheet
                    break
        
        if not target_sheet:
            print("일위대가 시트를 찾을 수 없습니다!")
            return None, None
        
        print(f"대상 시트: {target_sheet}")
        
        # 시트 읽기
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=None, engine='xlrd')
        else:
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
        
        # 구조 분석
        print("\n구조 분석:")
        
        # 1. 호표 패턴 확인
        hopyo_patterns = {
            'separated': False,  # 호표가 별도 열에 있는지
            'colon': False,      # 콜론 형식인지
            'parenthesis': False # 괄호 형식인지
        }
        
        # 처음 30행 검사
        for i in range(min(30, len(df))):
            for j in range(min(10, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell)
                    
                    # 호표 패턴 확인
                    if '호표' in cell_str:
                        if '：' in cell_str or ':' in cell_str:
                            hopyo_patterns['colon'] = True
                            print(f"  - 콜론 형식 호표 발견: {cell_str[:50]}")
                        elif '(' in cell_str and ')' in cell_str and '호표' in cell_str:
                            hopyo_patterns['parenthesis'] = True
                            print(f"  - 괄호 형식 호표 발견: {cell_str[:50]}")
                        elif re.match(r'^제?\s*\d+\s*호표$', cell_str.strip()):
                            hopyo_patterns['separated'] = True
                            print(f"  - 분리된 호표 발견: {cell_str}")
        
        # 2. 헤더 확인
        header_keywords = {
            '호표': False,
            '품명': False,
            '공종': False,
            '규격': False,
            '단위': False,
            '수량': False
        }
        
        for i in range(min(10, len(df))):
            for j in range(min(10, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    for keyword in header_keywords:
                        if keyword in cell_str:
                            header_keywords[keyword] = True
        
        print(f"\n헤더 키워드 발견: {[k for k, v in header_keywords.items() if v]}")
        
        # 3. 파일 타입 결정
        file_type = None
        
        # sgs 타입: 콜론 형식 + 공종 헤더
        if hopyo_patterns['colon']:
            file_type = 'sgs_type'
        # construction 타입: 괄호 형식
        elif hopyo_patterns['parenthesis']:
            file_type = 'construction_type'
        # test1, est 타입: 분리형
        elif hopyo_patterns['separated']:
            # est는 호표가 1열(인덱스 1)에, test1은 0열에
            for i in range(3, min(20, len(df))):
                if pd.notna(df.iloc[i, 1]) and '호표' in str(df.iloc[i, 1]):
                    file_type = 'est_type'
                    break
                elif pd.notna(df.iloc[i, 0]) and '호표' in str(df.iloc[i, 0]):
                    file_type = 'test1_type'
                    break
        
        if file_type:
            print(f"\n감지된 파일 타입: {self.file_patterns[file_type]['description']}")
            return file_type, target_sheet
        else:
            print("\n파일 타입을 자동으로 감지할 수 없습니다.")
            return None, None
    
    def parse_file(self, file_path: str) -> Dict[str, Any]:
        """파일을 자동으로 파싱"""
        
        # 파일 타입 감지
        file_type, sheet_name = self.detect_file_type(file_path)
        
        if not file_type:
            return {"error": "파일 구조를 인식할 수 없습니다"}
        
        # 적절한 파서 import 및 실행
        print(f"\n{file_type} 파서 실행 중...")
        
        if file_type == 'test1_type':
            from parser_test1 import parse_test1
            # 기존 파서는 고정 파일명을 사용하므로 직접 파싱
            result = self._parse_test1_type(file_path, sheet_name)
        elif file_type == 'sgs_type':
            from parser_sgs import parse_sgs
            result = self._parse_sgs_type(file_path, sheet_name)
        elif file_type == 'construction_type':
            from parser_construction import parse_construction
            result = self._parse_construction_type(file_path, sheet_name)
        elif file_type == 'est_type':
            from parser_est import parse_est
            result = self._parse_est_type(file_path, sheet_name)
        
        return result
    
    def _parse_test1_type(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """test1 타입 파싱"""
        # parser_test1.py의 로직 재사용
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        header_row = 1
        column_map = {
            '호표': 0,
            '품명': 1,
            '규격': 2,
            '단위': 3,
            '수량': 4,
            '비고': 11
        }
        
        hopyo_data = {}
        current_hopyo = None
        
        for row_idx in range(header_row + 1, len(df)):
            row = df.iloc[row_idx]
            
            # 호표 확인
            if '호표' in column_map:
                hopyo_cell = row[column_map['호표']]
                if pd.notna(hopyo_cell) and '호표' in str(hopyo_cell):
                    match = re.search(r'제?\s*(\d+)\s*호표', str(hopyo_cell))
                    if match:
                        hopyo_num = int(match.group(1))
                        current_hopyo = f"호표{hopyo_num}"
                        
                        work_name = str(row[column_map['품명']]).strip() if pd.notna(row[column_map['품명']]) else ""
                        work_spec = str(row[column_map['규격']]).strip() if pd.notna(row[column_map['규격']]) else ""
                        work_unit = str(row[column_map['단위']]).strip() if pd.notna(row[column_map['단위']]) else ""
                        
                        hopyo_data[current_hopyo] = {
                            "호표번호": hopyo_num,
                            "작업명": work_name,
                            "규격": work_spec,
                            "단위": work_unit,
                            "세부항목": []
                        }
                        continue
            
            # 세부 항목 추출
            if current_hopyo and '품명' in column_map:
                품명_cell = row[column_map['품명']]
                
                if pd.notna(품명_cell):
                    품명 = str(품명_cell).strip()
                    
                    if any(skip in 품명 for skip in ['합계', '소계', '재료비', '노무비', '경비']):
                        continue
                    
                    if not 품명:
                        continue
                    
                    item = {"품명": 품명}
                    item['규격'] = str(row[column_map['규격']]).strip() if pd.notna(row[column_map['규격']]) else ""
                    item['단위'] = str(row[column_map['단위']]).strip() if pd.notna(row[column_map['단위']]) else ""
                    
                    try:
                        item['수량'] = float(str(row[column_map['수량']]).replace(',', '')) if pd.notna(row[column_map['수량']]) else 0
                    except:
                        item['수량'] = 0
                    
                    item['비고'] = str(row[column_map['비고']]).strip() if pd.notna(row[column_map['비고']]) else ""
                    
                    hopyo_data[current_hopyo]['세부항목'].append(item)
        
        # 항목수 계산
        for data in hopyo_data.values():
            data['항목수'] = len(data['세부항목'])
        
        return {
            "file": file_path,
            "sheet": sheet_name,
            "file_type": "test1_type",
            "hopyo_count": len(hopyo_data),
            "hopyo_data": hopyo_data
        }
    
    def _parse_sgs_type(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """sgs 타입 파싱 - 간단히 구현"""
        # 실제 구현은 parser_sgs.py 참조
        return {"message": "sgs 타입 파서 구현 필요"}
    
    def _parse_construction_type(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """construction 타입 파싱 - 간단히 구현"""
        # 실제 구현은 parser_construction.py 참조
        return {"message": "construction 타입 파서 구현 필요"}
    
    def _parse_est_type(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """est 타입 파싱 - 간단히 구현"""
        # 실제 구현은 parser_est.py 참조
        return {"message": "est 타입 파서 구현 필요"}

def main():
    """메인 함수"""
    parser = AutoParser()
    
    # 테스트할 파일 경로
    test_files = [
        'test1.xlsx',
        'sgs.xls',
        '건축구조내역2.xlsx',
        'est.xlsx'
    ]
    
    print("="*80)
    print("자동 파일 구조 감지 및 파싱 테스트")
    print("="*80)
    
    for file_path in test_files:
        if os.path.exists(file_path):
            print(f"\n\n{'='*60}")
            print(f"파일: {file_path}")
            print(f"{'='*60}")
            
            result = parser.parse_file(file_path)
            
            if "error" not in result:
                output_file = f"{file_path.split('.')[0]}_auto.json"
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                
                print(f"\n결과 저장: {output_file}")
                if "hopyo_count" in result:
                    print(f"호표 수: {result['hopyo_count']}")
                    total_items = sum(data['항목수'] for data in result['hopyo_data'].values())
                    print(f"총 세부항목: {total_items}")
            else:
                print(f"\n오류: {result['error']}")

if __name__ == "__main__":
    main()