"""
test1.xlsx 일위대가_산근 시트 통합 파서
통합 JSON 구조: 품명, 규격, 단위, 수량 4개 컬럼으로 표준화
"""
import pandas as pd
import re
import sys
import io
from ilwidae_base_parser import IlwidaeBaseParser

# UTF-8 인코딩 설정
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except:
    pass

class IlwidaeTest1Parser(IlwidaeBaseParser):
    """test1.xlsx 일위대가 파서"""
    
    def __init__(self):
        super().__init__('test1.xlsx', '일위대가_산근')
        
    def read_list_info(self):
        """일위대가목록표 읽기"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='일위대가목록표', header=None)
            
            hopyo_info = []
            for i in range(len(df)):
                # 제N호표 패턴 찾기
                for j in range(len(df.columns)):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        # 제N호표 형식
                        match = re.match(r'제\s*(\d+)\s*호표', cell_str)
                        if match:
                            num = match.group(1)
                            
                            # 작업명 찾기 (다음 컬럼 또는 다음 행)
                            work = ''
                            if j+1 < len(df.columns):
                                work_cell = df.iloc[i, j+1]
                                if pd.notna(work_cell):
                                    work = str(work_cell).strip()
                            
                            # 규격, 단위 정보 수집
                            spec = ''
                            unit = ''
                            if j+2 < len(df.columns):
                                spec_cell = df.iloc[i, j+2]
                                if pd.notna(spec_cell):
                                    spec = str(spec_cell).strip()
                            if j+3 < len(df.columns):
                                unit_cell = df.iloc[i, j+3]
                                if pd.notna(unit_cell):
                                    unit = str(unit_cell).strip()
                            
                            hopyo_info.append({
                                'num': num,
                                'work': work,
                                'spec': spec,
                                'unit': unit
                            })
                            break
            
            print(f"\n[일위대가목록표 정보]")
            print(f"총 {len(hopyo_info)}개 일위대가 항목 발견")
            
            return hopyo_info
            
        except Exception as e:
            print(f"목록 읽기 오류: {str(e)}")
            return None
    
    def find_ilwidae_hopos(self):
        """일위대가 호표 찾기 (제N호표 패턴)"""
        hopyo_list = []
        
        # 엑셀 파일 읽기
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name, header=None)
            print(f"\n시트 '{self.sheet_name}' 로드 완료. 크기: {self.df.shape}")
        except Exception as e:
            print(f"파일 읽기 오류: {str(e)}")
            return []
        
        # 컬럼 감지
        self.detect_columns(self.df)
        print(f"컬럼 매핑: {self.column_map}")
        
        # 호표 패턴: 제N호표
        hopyo_pattern = r'제\s*(\d+)\s*호표'
        
        # 호표 찾기
        for i in range(len(self.df)):
            for j in range(min(5, len(self.df.columns))):  # 처음 5개 컬럼만 확인
                cell = self.df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    match = re.match(hopyo_pattern, cell_str)
                    if match:
                        hopyo_num = match.group(1)
                        
                        # 작업명 찾기
                        work_name = ''
                        # 같은 행의 다음 컬럼들 확인
                        for k in range(j+1, min(j+5, len(self.df.columns))):
                            next_cell = self.df.iloc[i, k]
                            if pd.notna(next_cell):
                                next_str = str(next_cell).strip()
                                if next_str and not next_str.isdigit():
                                    work_name = next_str
                                    break
                        
                        hopyo_list.append({
                            'row': i,
                            'num': hopyo_num,
                            'work': work_name,
                            'full_text': cell_str
                        })
                        break  # 한 행에서 호표 찾으면 다음 행으로
        
        return hopyo_list
    
    def parse_ilwidae_title(self, row_idx, df):
        """일위대가 타이틀 파싱 (test1 특화)"""
        title_data = {
            "품명": "",
            "규격": "",
            "단위": "",
            "수량": ""
        }
        
        # 호표 행에서 작업명 추출
        for j in range(min(10, len(df.columns))):
            cell = df.iloc[row_idx, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                # 제N호표 다음 내용이 작업명
                if '호표' in cell_str:
                    continue
                elif cell_str and not cell_str.isdigit() and not title_data["품명"]:
                    title_data["품명"] = cell_str
                elif cell_str in ['M3', 'M2', 'M', 'KG', 'TON', '개', '대', 'L', 'EA', 'M³', 'M²']:
                    title_data["단위"] = cell_str
                elif re.match(r'^\d+\.?\d*$', cell_str) and not title_data["수량"]:
                    title_data["수량"] = cell_str
        
        # 목록 정보에서 보완
        if self.list_info:
            for item in self.list_info:
                if item['work'] and (item['work'] == title_data["품명"] or title_data["품명"] in item['work']):
                    if not title_data["규격"] and item.get('spec'):
                        title_data["규격"] = item['spec']
                    if not title_data["단위"] and item.get('unit'):
                        title_data["단위"] = item['unit']
                    break
        
        return title_data

def main():
    """메인 실행 함수"""
    parser = IlwidaeTest1Parser()
    
    # 파싱 및 저장
    output_file = parser.save_to_json()
    
    if output_file:
        print(f"\n✅ test1.xlsx 일위대가 파싱 완료")
        print(f"   통합 구조 JSON 저장: {output_file}")

if __name__ == "__main__":
    main()