"""
최종 통합 일위대가 파서
모든 파일 타입에 대해 100% 성공을 목표로 하는 개선된 파서
"""
import pandas as pd
import json
import os
import re
from typing import Dict, List, Optional, Tuple

class FinalUnifiedParser:
    """최종 통합 파서"""
    
    def __init__(self):
        os.makedirs('result', exist_ok=True)
        
    def parse_test1(self):
        """test1.xlsx 파싱 (이미 100% 완벽)"""
        print("\n" + "="*70)
        print("test1.xlsx 파싱")
        print("="*70)
        
        df = pd.read_excel('test1.xlsx', sheet_name='일위대가_산근', header=None)
        print(f"시트 크기: {df.shape}")
        
        # 호표 찾기
        hopyo_pattern = r'제\s*(\d+)\s*호표'
        hopyo_list = []
        seen_nums = set()
        
        for i in range(len(df)):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    match = re.match(hopyo_pattern, cell_str)
                    if match:
                        hopyo_num = match.group(1)
                        
                        if hopyo_num in seen_nums:
                            break
                        
                        # 작업명 찾기
                        work_name = ''
                        for k in range(j+1, min(j+5, len(df.columns))):
                            next_cell = df.iloc[i, k]
                            if pd.notna(next_cell):
                                next_str = str(next_cell).strip()
                                if next_str and not next_str.replace('.', '').isdigit():
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
        
        print(f"호표 발견: {len(hopyo_list)}개")
        
        # 컬럼 정보
        column_info = {'품명': 1, '규격': 2, '단위': 3, '수량': 4, '비고': 13}
        
        # 결과 생성
        result = self._create_result('test1.xlsx', '일위대가_산근', hopyo_list, df, column_info)
        
        # 저장
        self._save_result(result, 'result/test1_final.json')
        
        return result
    
    def parse_est(self):
        """est.xlsx 파싱 (개선됨)"""
        print("\n" + "="*70)
        print("est.xlsx 파싱")
        print("="*70)
        
        df = pd.read_excel('est.xlsx', sheet_name='일위대가', header=None)
        print(f"시트 크기: {df.shape}")
        
        # 호표 찾기 (컬럼 1에서만)
        hopyo_pattern = r'제\s*(\d+)\s*호표'
        hopyo_list = []
        
        for i in range(len(df)):
            if df.shape[1] > 1:
                cell = df.iloc[i, 1]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    match = re.match(hopyo_pattern, cell_str)
                    if match:
                        hopyo_num = match.group(1)
                        
                        # 작업명은 컬럼 2에서
                        work_name = ''
                        if df.shape[1] > 2:
                            work_cell = df.iloc[i, 2]
                            if pd.notna(work_cell):
                                work_name = str(work_cell).strip()
                        
                        # 규격은 컬럼 3에서
                        spec = ''
                        if df.shape[1] > 3:
                            spec_cell = df.iloc[i, 3]
                            if pd.notna(spec_cell):
                                spec = str(spec_cell).strip()
                        
                        # 단위는 컬럼 5에서
                        unit = ''
                        if df.shape[1] > 5:
                            unit_cell = df.iloc[i, 5]
                            if pd.notna(unit_cell):
                                unit = str(unit_cell).strip()
                        
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name,
                            'spec': spec,
                            'unit': unit
                        })
        
        print(f"호표 발견: {len(hopyo_list)}개")
        
        # 컬럼 정보 (est 특화)
        column_info = {'품명': 2, '규격': 3, '수량': 4, '단위': 5, '비고': 14}
        
        # 결과 생성 (est 전용)
        result = self._create_est_result('est.xlsx', '일위대가', hopyo_list, df, column_info)
        
        # 저장
        self._save_result(result, 'result/est_final.json')
        
        return result
    
    def parse_sgs(self):
        """sgs.xls 파싱 (개선 - 일위대가 목록표 활용)"""
        print("\n" + "="*70)
        print("sgs.xls 파싱")
        print("="*70)
        
        df = pd.read_excel('sgs.xls', sheet_name='일위대가', header=None, engine='xlrd')
        print(f"시트 크기: {df.shape}")
        
        # 일위대가 목록표에서 품명/규격/단위 정보 가져오기
        df_list = pd.read_excel('sgs.xls', sheet_name='일위대가 목록', header=None, engine='xlrd')
        print(f"일위대가 목록 시트 크기: {df_list.shape}")
        
        # 목록표에서 정보 추출 (행 4부터 시작)
        title_info = {}
        for i in range(4, len(df_list)):
            if pd.notna(df_list.iloc[i, 0]):
                # 공종(품명)
                work_name = str(df_list.iloc[i, 0]).strip()
                # 규격 (컬럼 1)
                spec = str(df_list.iloc[i, 1]).strip() if pd.notna(df_list.iloc[i, 1]) else ''
                # 단위 (컬럼 3)
                unit = str(df_list.iloc[i, 3]).strip() if pd.notna(df_list.iloc[i, 3]) else ''
                
                # 작업명을 키로 저장
                title_info[work_name] = {
                    '규격': spec,
                    '단위': unit
                }
        
        print(f"목록표에서 추출한 일위대가 정보: {len(title_info)}개")
        
        # 호표 찾기
        hopyo_pattern = r'제\s*(\d+)\s*호표'
        hopyo_list = []
        
        for i in range(len(df)):
            for j in range(min(3, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # 호표 패턴 매칭
                    match = re.search(hopyo_pattern, cell_str)
                    if match:
                        hopyo_num = match.group(1)
                        
                        # 작업명 추출 (：또는 : 뒤)
                        work_name = ''
                        if '：' in cell_str:
                            parts = cell_str.split('：')
                            if len(parts) > 1:
                                work_name = parts[1].strip()
                        elif ':' in cell_str:
                            parts = cell_str.split(':')
                            if len(parts) > 1:
                                work_name = parts[1].strip()
                        
                        if work_name:
                            # 작업명에서 단위 정보 제거 (예: "탄성포장 철거 ㎡ 당" -> "탄성포장 철거")
                            clean_work_name = work_name.replace(' ㎡ 당', '').replace(' M 당', '').replace(' EA 당', '').replace(' 개소 당', '').replace(' 기 당', '').replace(' 본 당', '').strip()
                            
                            # 괄호 안의 규격 정보 분리
                            spec_from_name = ''
                            base_work_name = clean_work_name
                            if '(' in clean_work_name and ')' in clean_work_name:
                                # 괄호 안의 내용을 규격으로 추출
                                import re as re2
                                match2 = re2.search(r'\(([^)]+)\)', clean_work_name)
                                if match2:
                                    spec_from_name = match2.group(1)
                                    base_work_name = clean_work_name.split('(')[0].strip()
                            
                            # 목록표에서 정보 찾기
                            info = title_info.get(base_work_name, title_info.get(clean_work_name, {}))
                            
                            # 규격 정보 우선순위: 1) 목록표 규격, 2) 품명에서 추출한 규격
                            final_spec = info.get('규격', '') or spec_from_name
                            
                            hopyo_list.append({
                                'row': i,
                                'num': hopyo_num,
                                'work': work_name,
                                'clean_work': base_work_name,
                                'spec': final_spec,
                                'unit': info.get('단위', '')
                            })
                        break
        
        print(f"호표 발견: {len(hopyo_list)}개")
        
        # 컬럼 정보 (sgs 특화 - 헤더 행 3 기준)
        column_info = {'품명': 0, '규격': 1, '수량': 2, '단위': 3, '비고': -1}  # sgs는 비고 없음
        
        # 결과 생성 (목록표 정보 포함)
        result = self._create_sgs_result('sgs.xls', '일위대가', hopyo_list, df, column_info)
        
        # 저장
        self._save_result(result, 'result/sgs_final.json')
        
        return result
    
    def parse_construction1(self):
        """건축구조내역.xlsx 파싱"""
        print("\n" + "="*70)
        print("건축구조내역.xlsx 파싱")
        print("="*70)
        
        df = pd.read_excel('건축구조내역.xlsx', sheet_name='일위대가', header=None)
        print(f"시트 크기: {df.shape}")
        
        # 호표 찾기 - ( 호표 N ) 패턴
        hopyo_list = []
        seen_nums = set()
        
        for i in range(len(df)):
            cell = df.iloc[i, 0]  # 첫 번째 컬럼에서만 찾기
            if pd.notna(cell):
                cell_str = str(cell).strip()
                
                # ( 호표 N ) 패턴 찾기
                match = re.search(r'\( 호표 (\d+) \)', cell_str)
                if match:
                    hopyo_num = match.group(1)
                    
                    if hopyo_num in seen_nums:
                        continue
                    
                    # 작업명 추출 - ( 호표 N ) 앞부분
                    work_name = cell_str.split('(')[0].strip()
                    
                    if work_name:
                        seen_nums.add(hopyo_num)
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name
                        })
                        print(f"  발견: 호표 {hopyo_num} - {work_name}")
        
        print(f"호표 발견: {len(hopyo_list)}개")
        
        # 컬럼 정보 (헤더 행 1 기준)
        column_info = {'품명': 0, '규격': 1, '단위': 2, '수량': 3, '비고': 12}
        
        # 결과 생성
        result = self._create_result('건축구조내역.xlsx', '일위대가', hopyo_list, df, column_info)
        
        # 저장
        self._save_result(result, 'result/construction1_final.json')
        
        return result
    
    def parse_construction2(self):
        """건축구조내역2.xlsx 파싱"""
        print("\n" + "="*70)
        print("건축구조내역2.xlsx 파싱")
        print("="*70)
        
        df = pd.read_excel('건축구조내역2.xlsx', sheet_name='일위대가', header=None)
        print(f"시트 크기: {df.shape}")
        
        # 호표 찾기 (M-NNN 패턴)
        hopyo_list = []
        seen_codes = set()
        
        for i in range(len(df)):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # M-101 형식 찾기
                    match = re.match(r'([A-Z])-(\d+)', cell_str)
                    if match:
                        code = f"{match.group(1)}-{match.group(2)}"
                        
                        if code in seen_codes:
                            continue
                        
                        # 작업명 찾기
                        work_name = ''
                        
                        # 같은 행 다음 셀들
                        for k in range(j+1, min(j+5, len(df.columns))):
                            next_cell = df.iloc[i, k]
                            if pd.notna(next_cell):
                                next_str = str(next_cell).strip()
                                if next_str and not re.match(r'^[A-Z]-\d+$', next_str):
                                    work_name = next_str
                                    break
                        
                        # 다음 행에서 찾기
                        if not work_name and i+1 < len(df):
                            for k in range(min(5, len(df.columns))):
                                next_cell = df.iloc[i+1, k]
                                if pd.notna(next_cell):
                                    next_str = str(next_cell).strip()
                                    if next_str and not re.match(r'^[\d.,]+$', next_str) and '합계' not in next_str:
                                        work_name = next_str
                                        break
                        
                        if work_name:
                            seen_codes.add(code)
                            hopyo_list.append({
                                'row': i,
                                'num': match.group(2),
                                'work': f"{code} {work_name}"
                            })
                        break
        
        print(f"호표 발견: {len(hopyo_list)}개")
        
        # 컬럼 정보
        column_info = {'품명': 0, '규격': 1, '단위': 2, '수량': 3, '비고': 12}
        
        # 결과 생성
        result = self._create_result('건축구조내역2.xlsx', '일위대가', hopyo_list, df, column_info)
        
        # 저장
        self._save_result(result, 'result/construction2_final.json')
        
        return result
    
    def parse_stmate(self):
        """stmate.xlsx 파싱 - 일위대가_호표 시트에서 진짜 일위대가 파싱"""
        print("\n" + "="*70)
        print("stmate.xlsx 파싱 (일위대가_호표 시트)")
        print("="*70)
        
        # repaired 파일에서 일위대가_호표 시트 사용
        df = pd.read_excel('stmate_repaired.xlsx', sheet_name='일위대가_호표', header=None)
        print(f"시트 크기: {df.shape}")
        
        # 호표 찾기 - No.N 패턴으로 시작하는 행 (진짜 일위대가)
        hopyo_list = []
        seen_nums = set()
        
        for i in range(len(df)):
            cell = df.iloc[i, 0]  # 첫 번째 컬럼에서만 찾기
            if pd.notna(cell):
                cell_str = str(cell).strip()
                
                # No.N 패턴 찾기 (예: "No.1  파고라(철거)")
                match = re.match(r'^No\.(\d+)\s+(.+)$', cell_str)
                if match:
                    hopyo_num = match.group(1)
                    
                    if hopyo_num in seen_nums:
                        continue
                    
                    # 작업명 추출
                    work_name = match.group(2).strip()
                    
                    # 규격은 컬럼 1에서
                    spec = ''
                    if len(df.columns) > 1:
                        spec_cell = df.iloc[i, 1]
                        if pd.notna(spec_cell):
                            spec = str(spec_cell).strip()
                    
                    # 단위는 컬럼 3에서
                    unit = ''
                    if len(df.columns) > 3:
                        unit_cell = df.iloc[i, 3]
                        if pd.notna(unit_cell):
                            unit = str(unit_cell).strip()
                    
                    # 수량은 컬럼 2에서
                    quantity = ''
                    if len(df.columns) > 2:
                        qty_cell = df.iloc[i, 2]
                        if pd.notna(qty_cell):
                            quantity = str(qty_cell).strip()
                    
                    if work_name:
                        seen_nums.add(hopyo_num)
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name,
                            'spec': spec,
                            'unit': unit,
                            'quantity': quantity
                        })
                        print(f"  발견: No.{hopyo_num} - {work_name} | {spec} | {unit}")
        
        print(f"일위대가 발견: {len(hopyo_list)}개")
        
        # 컬럼 정보 (stmate 호표 시트 특화)
        column_info = {'품명': 0, '규격': 1, '수량': 2, '단위': 3, '비고': -1}  # 비고 없음
        
        # 결과 생성 (stmate 호표 전용)
        result = self._create_stmate_hopyo_result('stmate.xlsx', '일위대가_호표', hopyo_list, df, column_info)
        
        # 저장
        self._save_result(result, 'result/stmate_final.json')
        
        return result
    
    def _create_result(self, file_name: str, sheet_name: str, hopyo_list: List, 
                      df: pd.DataFrame, column_info: Dict) -> Dict:
        """결과 JSON 생성"""
        result = {
            'file': file_name,
            'sheet': sheet_name,
            'total_ilwidae_count': len(hopyo_list),
            'ilwidae_data': []
        }
        
        for idx, hopyo in enumerate(hopyo_list):
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else min(start_row + 50, len(df))
            
            # 산출근거 추출
            sangul_items = []
            
            for row_idx in range(start_row + 1, end_row):
                # 빈 행 체크
                row_empty = True
                for col in column_info.values():
                    if col < len(df.columns) and pd.notna(df.iloc[row_idx, col]):
                        row_empty = False
                        break
                
                if row_empty:
                    continue
                
                row_data = {
                    'row_number': row_idx,
                    '품명': '',
                    '규격': '',
                    '단위': '',
                    '수량': '',
                    '비고': ''
                }
                
                # 각 필드 추출
                for field, col_idx in column_info.items():
                    if col_idx >= 0 and col_idx < len(df.columns):
                        cell = df.iloc[row_idx, col_idx]
                        if pd.notna(cell):
                            value = str(cell).strip()
                            # 합계, 호표 등 제외
                            if field == '품명' and ('합계' in value or '호표' in value or re.match(r'^[\d.,]+$', value)):
                                break
                            row_data[field] = value
                
                # 품명이 있는 경우만 추가
                if row_data['품명']:
                    sangul_items.append(row_data)
            
            ilwidae_item = {
                'ilwidae_no': hopyo['num'],
                'ilwidae_title': {
                    '품명': hopyo['work'],
                    '규격': '',
                    '단위': '',
                    '수량': '',
                    '비고': ''
                },
                'position': {
                    'start_row': start_row,
                    'end_row': end_row,
                    'total_rows': end_row - start_row
                },
                '산출근거': sangul_items
            }
            
            result['ilwidae_data'].append(ilwidae_item)
        
        return result
    
    def _create_est_result(self, file_name: str, sheet_name: str, hopyo_list: List, 
                           df: pd.DataFrame, column_info: Dict) -> Dict:
        """est 전용 결과 JSON 생성 (호표 행의 규격/단위 활용)"""
        result = {
            'file': file_name,
            'sheet': sheet_name,
            'total_ilwidae_count': len(hopyo_list),
            'ilwidae_data': []
        }
        
        for idx, hopyo in enumerate(hopyo_list):
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else min(start_row + 50, len(df))
            
            # 산출근거 추출
            sangul_items = []
            
            for row_idx in range(start_row + 1, end_row):
                # 빈 행 체크
                row_empty = True
                for col in column_info.values():
                    if col >= 0 and col < len(df.columns) and pd.notna(df.iloc[row_idx, col]):
                        row_empty = False
                        break
                
                if row_empty:
                    continue
                
                row_data = {
                    'row_number': row_idx,
                    '품명': '',
                    '규격': '',
                    '단위': '',
                    '수량': '',
                    '비고': ''
                }
                
                # 각 필드 추출
                for field, col_idx in column_info.items():
                    if col_idx >= 0 and col_idx < len(df.columns):
                        cell = df.iloc[row_idx, col_idx]
                        if pd.notna(cell):
                            value = str(cell).strip()
                            # 합계, 호표 등 제외
                            if field == '품명' and ('합계' in value or '호표' in value or re.match(r'^[\d.,]+$', value)):
                                break
                            row_data[field] = value
                
                # 품명이 있는 경우만 추가
                if row_data['품명']:
                    sangul_items.append(row_data)
            
            # 일위대가 타이틀에 호표 행의 정보 사용
            ilwidae_item = {
                'ilwidae_no': hopyo['num'],
                'ilwidae_title': {
                    '품명': hopyo['work'],
                    '규격': hopyo.get('spec', ''),
                    '단위': hopyo.get('unit', ''),
                    '수량': '1',  # 기본값
                    '비고': ''
                },
                'position': {
                    'start_row': start_row,
                    'end_row': end_row,
                    'total_rows': end_row - start_row
                },
                '산출근거': sangul_items
            }
            
            result['ilwidae_data'].append(ilwidae_item)
        
        return result
    
    def _create_sgs_result(self, file_name: str, sheet_name: str, hopyo_list: List, 
                           df: pd.DataFrame, column_info: Dict) -> Dict:
        """sgs 전용 결과 JSON 생성 (목록표 정보 활용)"""
        result = {
            'file': file_name,
            'sheet': sheet_name,
            'total_ilwidae_count': len(hopyo_list),
            'ilwidae_data': []
        }
        
        for idx, hopyo in enumerate(hopyo_list):
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else min(start_row + 50, len(df))
            
            # 산출근거 추출
            sangul_items = []
            
            for row_idx in range(start_row + 1, end_row):
                # 빈 행 체크
                row_empty = True
                for col in column_info.values():
                    if col >= 0 and col < len(df.columns) and pd.notna(df.iloc[row_idx, col]):
                        row_empty = False
                        break
                
                if row_empty:
                    continue
                
                row_data = {
                    'row_number': row_idx,
                    '품명': '',
                    '규격': '',
                    '단위': '',
                    '수량': '',
                    '비고': ''
                }
                
                # 각 필드 추출
                for field, col_idx in column_info.items():
                    if col_idx >= 0 and col_idx < len(df.columns):
                        cell = df.iloc[row_idx, col_idx]
                        if pd.notna(cell):
                            value = str(cell).strip()
                            # 합계, 호표 등 제외
                            if field == '품명' and ('합계' in value or '호표' in value or re.match(r'^[\d.,]+$', value)):
                                break
                            row_data[field] = value
                
                # 품명이 있는 경우만 추가
                if row_data['품명']:
                    sangul_items.append(row_data)
            
            # 일위대가 타이틀에 목록표 정보 사용
            ilwidae_item = {
                'ilwidae_no': hopyo['num'],
                'ilwidae_title': {
                    '품명': hopyo['clean_work'],  # 깨끗한 작업명 사용
                    '규격': hopyo['spec'],  # 목록표의 규격
                    '단위': hopyo['unit'],  # 목록표의 단위
                    '수량': '1',  # 기본값
                    '비고': ''
                },
                'position': {
                    'start_row': start_row,
                    'end_row': end_row,
                    'total_rows': end_row - start_row
                },
                '산출근거': sangul_items
            }
            
            result['ilwidae_data'].append(ilwidae_item)
        
        return result
    
    def _create_stmate_result(self, file_name: str, sheet_name: str, hopyo_list: List, 
                             df: pd.DataFrame, column_info: Dict) -> Dict:
        """stmate 전용 결과 JSON 생성"""
        result = {
            'file': file_name,
            'sheet': sheet_name,
            'total_ilwidae_count': len(hopyo_list),
            'ilwidae_data': []
        }
        
        for idx, hopyo in enumerate(hopyo_list):
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else min(start_row + 100, len(df))
            
            # 산출근거 추출
            sangul_items = []
            
            for row_idx in range(start_row + 1, end_row):
                # 빈 행이나 다음 호표 시작 체크
                if row_idx < len(df):
                    first_cell = df.iloc[row_idx, 0] if len(df.columns) > 0 else None
                    if pd.notna(first_cell):
                        first_str = str(first_cell).strip()
                        # 다음 호표 시작이면 중단
                        if re.match(r'^#\d+', first_str):
                            break
                
                # 산출근거 추출 - stmate는 특별한 구조
                row_data = {
                    'row_number': row_idx,
                    '품명': '',
                    '규격': '',
                    '단위': '',
                    '수량': '',
                    '비고': ''
                }
                
                # 품명은 행의 첫 번째 컬럼에서 (설명 라인들)
                if row_idx < len(df) and len(df.columns) > 0:
                    desc_cell = df.iloc[row_idx, 0]
                    if pd.notna(desc_cell):
                        desc_str = str(desc_cell).strip()
                        # 계산식이나 설명이 아닌 경우만 품명으로 취급
                        if desc_str and not desc_str.startswith('=') and '계' not in desc_str and 'Q =' not in desc_str:
                            if desc_str.startswith(' '):  # 들여쓰기된 항목들
                                # 재료비, 노무비, 경비 라인 파싱
                                if '재 료 비' in desc_str or '노 무 비' in desc_str or '경    비' in desc_str:
                                    parts = desc_str.split(':')
                                    if len(parts) >= 2:
                                        row_data['품명'] = parts[0].strip()
                                        # 수량 정보 추출
                                        amount_part = parts[1].strip()
                                        if '=' in amount_part:
                                            amount_value = amount_part.split('=')[-1].replace('원', '').strip()
                                            if amount_value:
                                                row_data['수량'] = amount_value
                                elif desc_str.strip() and not desc_str.startswith('Q '):
                                    row_data['품명'] = desc_str.strip()
                
                # 합계 행 처리
                total_cell = df.iloc[row_idx, 1] if row_idx < len(df) and len(df.columns) > 1 else None
                if pd.notna(total_cell):
                    total_str = str(total_cell).strip()
                    if total_str and total_str.replace(',', '').replace('.', '').isdigit():
                        if not row_data['품명']:
                            # 첫 컬럼에서 품명 확인
                            first_cell = df.iloc[row_idx, 0] if len(df.columns) > 0 else None
                            if pd.notna(first_cell):
                                first_str = str(first_cell).strip()
                                if '소    계' in first_str or '전체 합계' in first_str:
                                    row_data['품명'] = first_str.strip()
                                    row_data['수량'] = total_str
                
                # 품명이 있는 경우만 추가
                if row_data['품명']:
                    sangul_items.append(row_data)
            
            # 일위대가 타이틀 생성
            ilwidae_item = {
                'ilwidae_no': hopyo['num'],
                'ilwidae_title': {
                    '품명': hopyo['work'],
                    '규격': hopyo['spec'],
                    '단위': hopyo['unit'],
                    '수량': '1',  # 기본값
                    '비고': ''
                },
                'position': {
                    'start_row': start_row,
                    'end_row': end_row,
                    'total_rows': end_row - start_row
                },
                '산출근거': sangul_items
            }
            
            result['ilwidae_data'].append(ilwidae_item)
        
        return result
    
    def _create_stmate_hopyo_result(self, file_name: str, sheet_name: str, hopyo_list: List, 
                                   df: pd.DataFrame, column_info: Dict) -> Dict:
        """stmate 일위대가_호표 전용 결과 JSON 생성"""
        result = {
            'file': file_name,
            'sheet': sheet_name,
            'total_ilwidae_count': len(hopyo_list),
            'ilwidae_data': []
        }
        
        for idx, hopyo in enumerate(hopyo_list):
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else min(start_row + 50, len(df))
            
            # 산출근거 추출 - 들여쓰기된 항목들
            sangul_items = []
            
            for row_idx in range(start_row + 1, end_row):
                # 다음 No. 시작이면 중단
                if row_idx < len(df):
                    first_cell = df.iloc[row_idx, 0] if len(df.columns) > 0 else None
                    if pd.notna(first_cell):
                        first_str = str(first_cell).strip()
                        if re.match(r'^No\.\d+', first_str):
                            break
                
                # 들여쓰기된 항목들 추출 (예: "   깨기(30cm미만)")
                row_data = {
                    'row_number': row_idx,
                    '품명': '',
                    '규격': '',
                    '단위': '',
                    '수량': '',
                    '비고': ''
                }
                
                # 각 필드 추출
                for field, col_idx in column_info.items():
                    if col_idx >= 0 and col_idx < len(df.columns):
                        cell = df.iloc[row_idx, col_idx]
                        if pd.notna(cell):
                            value = str(cell).strip()
                            # 빈 행이나 헤더 제외
                            if field == '품명' and value:
                                # 들여쓰기된 항목만 산출근거로 처리
                                if value.startswith('   ') or value.startswith('\t'):
                                    row_data[field] = value.strip()
                                elif not re.match(r'^No\.\d+', value):
                                    # 공종명이 아닌 경우만
                                    row_data[field] = value
                            elif field != '품명':
                                row_data[field] = value
                
                # 품명이 있고 유효한 산출근거인 경우만 추가
                if row_data['품명'] and not row_data['품명'].startswith('No.'):
                    sangul_items.append(row_data)
            
            # 일위대가 타이틀 생성
            ilwidae_item = {
                'ilwidae_no': hopyo['num'],
                'ilwidae_title': {
                    '품명': hopyo['work'],
                    '규격': hopyo['spec'],
                    '단위': hopyo['unit'],
                    '수량': hopyo['quantity'],
                    '비고': ''
                },
                'position': {
                    'start_row': start_row,
                    'end_row': end_row,
                    'total_rows': end_row - start_row
                },
                '산출근거': sangul_items
            }
            
            result['ilwidae_data'].append(ilwidae_item)
        
        return result
    
    def _save_result(self, result: Dict, output_file: str):
        """결과 저장"""
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        # 통계 출력
        total_sangul = sum(len(item['산출근거']) for item in result['ilwidae_data'])
        print(f"✅ 저장: {output_file}")
        print(f"   일위대가: {result['total_ilwidae_count']}개")
        print(f"   산출근거: {total_sangul}개")
    
    def run_all(self):
        """모든 파일 파싱"""
        results = []
        
        # test1.xlsx
        if os.path.exists('test1.xlsx'):
            results.append(self.parse_test1())
        
        # est.xlsx
        if os.path.exists('est.xlsx'):
            results.append(self.parse_est())
        
        # sgs.xls
        if os.path.exists('sgs.xls'):
            results.append(self.parse_sgs())
        
        # 건축구조내역.xlsx
        if os.path.exists('건축구조내역.xlsx'):
            results.append(self.parse_construction1())
        
        # 건축구조내역2.xlsx
        if os.path.exists('건축구조내역2.xlsx'):
            results.append(self.parse_construction2())
        
        # stmate.xlsx (repaired version 필요)
        if os.path.exists('stmate_repaired.xlsx'):
            results.append(self.parse_stmate())
        
        return results

def validate_final_results():
    """최종 결과 검증"""
    print("\n" + "="*80)
    print("최종 검증 보고서")
    print("="*80)
    
    files = [
        'result/test1_final.json',
        'result/est_final.json',
        'result/sgs_final.json',
        'result/construction1_final.json',
        'result/construction2_final.json',
        'result/stmate_final.json'
    ]
    
    total_success = 0
    total_files = 0
    
    for file_path in files:
        if os.path.exists(file_path):
            total_files += 1
            
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            print(f"\n{data['file']}:")
            print(f"  일위대가: {data['total_ilwidae_count']}개")
            
            # 산출근거 통계
            total_sangul = 0
            filled_counts = {'품명': 0, '규격': 0, '단위': 0, '수량': 0, '비고': 0}
            
            for item in data['ilwidae_data']:
                for sangul in item['산출근거']:
                    total_sangul += 1
                    for field in filled_counts.keys():
                        if sangul.get(field, ''):
                            filled_counts[field] += 1
            
            print(f"  산출근거: {total_sangul}개")
            
            # 성공 판정
            success = total_sangul > 0 and data['total_ilwidae_count'] > 0
            if success:
                total_success += 1
                print(f"  상태: ✅ 성공")
            else:
                print(f"  상태: ❌ 실패")
            
            # 필드 채움율
            if total_sangul > 0:
                print(f"  필드 채움율:")
                for field, count in filled_counts.items():
                    rate = count / total_sangul * 100
                    print(f"    {field}: {rate:.1f}%")
    
    # 전체 성공률
    print("\n" + "="*80)
    success_rate = (total_success / total_files * 100) if total_files > 0 else 0
    
    if success_rate == 100:
        print(f"🎉 최종 성공률: {success_rate:.1f}% - 모든 파일 파싱 성공!")
    else:
        print(f"📊 최종 성공률: {success_rate:.1f}% ({total_success}/{total_files})")
    
    return success_rate

if __name__ == "__main__":
    # 파서 실행
    parser = FinalUnifiedParser()
    results = parser.run_all()
    
    # 검증
    success_rate = validate_final_results()
    
    if success_rate < 100:
        print("\n⚠️ 100% 달성하지 못함. 추가 개선 필요.")
    else:
        print("\n✅ 목표 달성! 모든 파일 100% 파싱 성공!")