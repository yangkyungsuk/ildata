"""
일위대가 통합 파서 베이스 클래스
모든 일위대가 파서가 상속받아 사용하는 기본 클래스
통합 JSON 구조: 품명, 규격, 단위, 수량 4개 컬럼으로 표준화
"""
import pandas as pd
import json
import re
import sys
import io
from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Tuple

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

class IlwidaeBaseParser(ABC):
    """일위대가 파서 베이스 클래스"""
    
    def __init__(self, file_path: str, sheet_name: str):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.column_map = {}
        self.df = None
        self.list_info = None
        
    def detect_columns(self, df: pd.DataFrame, start_row: int = 0, end_row: int = 30) -> Dict[str, int]:
        """품명, 규격, 단위, 수량 컬럼 자동 감지"""
        column_patterns = {
            "품명": ["품명", "품 명", "자재명", "명칭", "공종", "품목"],
            "규격": ["규격", "규 격", "사양", "규", "SPEC"],
            "단위": ["단위", "단 위", "UNIT", "단"],
            "수량": ["수량", "수 량", "물량", "QTY", "량"]
        }
        
        column_map = {
            "품명": None,
            "규격": None,
            "단위": None,
            "수량": None
        }
        
        # 헤더 후보 행들 검사
        for row_idx in range(start_row, min(end_row, len(df))):
            row_data = []
            for col_idx in range(min(20, len(df.columns))):
                cell = df.iloc[row_idx, col_idx]
                if pd.notna(cell):
                    row_data.append(str(cell).strip())
                else:
                    row_data.append("")
            
            # 각 컬럼 패턴 매칭
            for col_name, patterns in column_patterns.items():
                if column_map[col_name] is None:  # 아직 찾지 못한 컬럼만
                    for col_idx, cell_value in enumerate(row_data):
                        for pattern in patterns:
                            if pattern in cell_value:
                                column_map[col_name] = col_idx
                                break
                        if column_map[col_name] is not None:
                            break
            
            # 모든 컬럼을 찾았으면 종료
            if all(v is not None for v in column_map.values()):
                break
        
        # 기본값 설정 (찾지 못한 컬럼)
        if column_map["품명"] is None:
            column_map["품명"] = 0  # 첫 번째 컬럼을 품명으로 가정
        if column_map["규격"] is None:
            column_map["규격"] = 1  # 두 번째 컬럼을 규격으로 가정
        if column_map["단위"] is None:
            column_map["단위"] = 2  # 세 번째 컬럼을 단위로 가정
        if column_map["수량"] is None:
            column_map["수량"] = 3  # 네 번째 컬럼을 수량으로 가정
        
        self.column_map = column_map
        return column_map
    
    def parse_ilwidae_title(self, row_idx: int, df: pd.DataFrame) -> Dict:
        """일위대가 타이틀 파싱 (4개 컬럼)"""
        title_data = {
            "품명": "",
            "규격": "",
            "단위": "",
            "수량": ""
        }
        
        # 호표 행에서 작업명 추출
        if self.column_map["품명"] is not None:
            cell = df.iloc[row_idx, self.column_map["품명"]]
            if pd.notna(cell):
                # 호표 패턴 제거하고 작업명만 추출
                work_name = re.sub(r'^(#\d+|No\.\d+|제\d+호표)\s*', '', str(cell).strip())
                title_data["품명"] = work_name
        
        # 규격, 단위, 수량 찾기 (같은 행 또는 다음 행에서)
        for col_name in ["규격", "단위", "수량"]:
            if self.column_map[col_name] is not None:
                # 같은 행에서 찾기
                cell = df.iloc[row_idx, self.column_map[col_name]]
                if pd.notna(cell) and str(cell).strip():
                    title_data[col_name] = str(cell).strip()
                # 다음 행에서 찾기
                elif row_idx + 1 < len(df):
                    cell = df.iloc[row_idx + 1, self.column_map[col_name]]
                    if pd.notna(cell) and str(cell).strip():
                        title_data[col_name] = str(cell).strip()
        
        return title_data
    
    def parse_sangul_items(self, df: pd.DataFrame, start_row: int, end_row: int) -> List[Dict]:
        """산출근거 항목들 파싱 (4개 컬럼)"""
        items = []
        
        for row_idx in range(start_row + 1, end_row):  # 타이틀 다음 행부터
            item_data = {
                "row_number": row_idx,
                "품명": "",
                "규격": "",
                "단위": "",
                "수량": "",
                "비고": None
            }
            
            # 각 컬럼 데이터 추출
            has_content = False
            for col_name in ["품명", "규격", "단위", "수량"]:
                if self.column_map[col_name] is not None:
                    cell = df.iloc[row_idx, self.column_map[col_name]]
                    if pd.notna(cell):
                        value = str(cell).strip()
                        if value:
                            item_data[col_name] = value
                            has_content = True
            
            # 비고 컬럼 확인 (수량 다음 컬럼)
            if self.column_map["수량"] is not None:
                remark_col = self.column_map["수량"] + 1
                if remark_col < len(df.columns):
                    cell = df.iloc[row_idx, remark_col]
                    if pd.notna(cell) and str(cell).strip():
                        item_data["비고"] = str(cell).strip()
            
            # 내용이 있는 행만 추가
            if has_content:
                items.append(item_data)
        
        return items
    
    def extract_raw_data(self, df: pd.DataFrame, start_row: int, end_row: int) -> List[Dict]:
        """기존 형식 호환을 위한 raw 데이터 추출"""
        raw_data = []
        
        for row_idx in range(start_row, end_row):
            row_data = {
                'row_number': row_idx,
                'columns': {},
                'has_content': False
            }
            
            # 모든 컬럼 데이터 수집
            for col_idx in range(min(20, len(df.columns))):
                cell = df.iloc[row_idx, col_idx]
                if pd.notna(cell):
                    content = str(cell).strip()
                    if content:
                        row_data['columns'][f'col_{col_idx}'] = content
                        row_data['has_content'] = True
            
            raw_data.append(row_data)
        
        return raw_data
    
    @abstractmethod
    def read_list_info(self) -> Optional[List[Dict]]:
        """목록표 읽기 (파서별 구현 필요)"""
        pass
    
    @abstractmethod
    def find_ilwidae_hopos(self) -> List[Dict]:
        """일위대가 호표 찾기 (파서별 구현 필요)"""
        pass
    
    def to_unified_json(self) -> Dict:
        """통합 JSON 구조로 변환"""
        if self.df is None:
            return None
        
        # 목록 정보 읽기
        self.list_info = self.read_list_info()
        
        # 컬럼 매핑
        self.detect_columns(self.df)
        
        # 일위대가 호표 찾기
        hopyo_list = self.find_ilwidae_hopos()
        
        print(f"\n[파싱 결과]")
        print(f"파일: {self.file_path}")
        print(f"시트: {self.sheet_name}")
        print(f"총 {len(hopyo_list)}개 일위대가 발견")
        
        # 검증
        if self.list_info:
            print(f"\n[검증 결과]")
            print(f"목록표: {len(self.list_info)}개, 실제 파싱: {len(hopyo_list)}개")
            match = len(self.list_info) == len(hopyo_list)
            print("✓ 일위대가 개수 일치" if match else "✗ 일위대가 개수 불일치")
        
        # 결과 구성
        result = {
            'file': self.file_path,
            'sheet': self.sheet_name,
            'total_ilwidae_count': len(hopyo_list),
            'validation': {
                'list_count': len(self.list_info) if self.list_info else 0,
                'parsed_count': len(hopyo_list),
                'match': len(self.list_info) == len(hopyo_list) if self.list_info else False
            },
            'ilwidae_data': []
        }
        
        # 각 일위대가 데이터 구성
        for idx, hopyo in enumerate(hopyo_list):
            # 범위 설정
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else len(self.df)
            
            # 타이틀 파싱
            title_data = self.parse_ilwidae_title(hopyo['row'], self.df)
            
            # 산출근거 파싱
            sangul_items = self.parse_sangul_items(self.df, start_row, end_row)
            
            # Raw 데이터 (호환성)
            raw_data = self.extract_raw_data(self.df, start_row, end_row)
            
            ilwidae_item = {
                'ilwidae_no': hopyo['num'],
                'ilwidae_title': title_data,
                'position': {
                    'start_row': start_row,
                    'end_row': end_row,
                    'total_rows': end_row - start_row
                },
                '산출근거': sangul_items,
                'raw_data': raw_data  # 기존 형식 호환
            }
            
            result['ilwidae_data'].append(ilwidae_item)
            
            print(f"\n일위대가 {hopyo['num']}: {title_data['품명']}")
            print(f"  - 위치: 행 {start_row} ~ {end_row-1}")
            print(f"  - 산출근거 항목: {len(sangul_items)}개")
        
        return result
    
    def save_to_json(self, output_file: str = None):
        """JSON 파일로 저장"""
        result = self.to_unified_json()
        if result is None:
            return None
        
        if output_file is None:
            # 기본 출력 파일명
            base_name = self.file_path.split('.')[0]
            sheet_name = self.sheet_name.replace(' ', '_').replace('/', '_')
            output_file = f"{base_name}_{sheet_name}_unified.json"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n✅ 통합 구조로 저장 완료: {output_file}")
        return output_file