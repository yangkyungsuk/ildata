"""
ebs.xls 일위대가_산근 시트 통합 파서
통합 JSON 구조: 품명, 규격, 단위, 수량 4개 컬럼으로 표준화
"""
import pandas as pd
import re
import sys
import io
from ilwidae_base_parser import IlwidaeBaseParser

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

class IlwidaeEbsParser(IlwidaeBaseParser):
    """ebs.xls 일위대가 파서"""
    
    def __init__(self):
        super().__init__('ebs.xls', '일위대가_산근')
        
    def read_list_info(self):
        """일위대가목록 읽기"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='일위대가목록', header=None, engine='xlrd')
            
            hopyo_info = []
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        # #N 형식의 일위대가 항목 찾기
                        match = re.match(r'^#(\d+)\s+(.+)', cell_str)
                        if match:
                            num = match.group(1)
                            work = match.group(2)
                            
                            # 규격, 단위 정보 수집
                            spec = df.iloc[i, j+1] if j+1 < len(df.columns) else ''
                            unit = df.iloc[i, j+2] if j+2 < len(df.columns) else ''
                            
                            hopyo_info.append({
                                'num': num,
                                'work': work,
                                'spec': str(spec).strip() if pd.notna(spec) else '',
                                'unit': str(unit).strip() if pd.notna(unit) else ''
                            })
                            break
            
            print(f"\n[일위대가목록 정보]")
            print(f"총 {len(hopyo_info)}개 일위대가 항목 발견")
            
            return hopyo_info
            
        except Exception as e:
            print(f"목록 읽기 오류: {str(e)}")
            return None
    
    def find_ilwidae_hopos(self):
        """일위대가 호표 찾기 (#N 패턴)"""
        hopyo_list = []
        
        # 엑셀 파일 읽기
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name, header=None, engine='xlrd')
            print(f"\n시트 '{self.sheet_name}' 로드 완료. 크기: {self.df.shape}")
        except Exception as e:
            print(f"파일 읽기 오류: {str(e)}")
            return []
        
        # 헤더 찾기 (산출근거 컬럼 찾기)
        header_row = None
        sangcul_col = 0
        
        for i in range(min(20, len(self.df))):
            for j in range(min(5, len(self.df.columns))):
                cell = self.df.iloc[i, j]
                if pd.notna(cell) and '산출근거' in str(cell):
                    header_row = i
                    sangcul_col = j
                    print(f"헤더 발견: 행 {header_row}, '산출근거' 컬럼 {sangcul_col}")
                    break
            if header_row is not None:
                break
        
        # 컬럼 감지 (헤더 행 기준)
        if header_row is not None:
            self.detect_columns(self.df, header_row, header_row + 5)
            print(f"컬럼 매핑: {self.column_map}")
        
        # 호표 패턴: #N 작업명 (일위대가용)
        hopyo_pattern = r'^#(\d+)\s+(.+)'
        
        # 호표 찾기
        start_row = header_row + 1 if header_row else 0
        for i in range(start_row, len(self.df)):
            cell = self.df.iloc[i, sangcul_col]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(hopyo_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    work_name = match.group(2).strip()
                    
                    hopyo_list.append({
                        'row': i,
                        'num': hopyo_num,
                        'work': work_name,
                        'full_text': cell_str
                    })
        
        return hopyo_list
    
    def parse_ilwidae_title(self, row_idx, df):
        """일위대가 타이틀 파싱 (ebs 특화)"""
        title_data = {
            "품명": "",
            "규격": "",
            "단위": "",
            "수량": ""
        }
        
        # 호표 행에서 작업명 추출
        cell = df.iloc[row_idx, 0]  # 첫 번째 컬럼 (산출근거)
        if pd.notna(cell):
            # #N 패턴 제거하고 작업명만 추출
            work_name = re.sub(r'^#\d+\s+', '', str(cell).strip())
            title_data["품명"] = work_name
        
        # 목록 정보에서 규격, 단위 가져오기
        if self.list_info:
            for item in self.list_info:
                if item['work'] == work_name or work_name in item['work']:
                    title_data["규격"] = item.get('spec', '')
                    title_data["단위"] = item.get('unit', '')
                    break
        
        # 같은 행 또는 다음 행에서 추가 정보 찾기
        for col_idx in range(1, min(10, len(df.columns))):
            cell = df.iloc[row_idx, col_idx]
            if pd.notna(cell):
                value = str(cell).strip()
                # 단위 패턴 체크
                if not title_data["단위"] and value in ['M3', 'M2', 'M', 'KG', 'TON', '개', '대', 'L', 'EA']:
                    title_data["단위"] = value
                # 숫자 패턴 체크 (수량)
                elif not title_data["수량"] and re.match(r'^\d+\.?\d*$', value):
                    title_data["수량"] = value
        
        return title_data

def main():
    """메인 실행 함수"""
    parser = IlwidaeEbsParser()
    
    # 파싱 및 저장
    output_file = parser.save_to_json('ebs_ilwidae_unified.json')
    
    if output_file:
        print(f"\n✅ ebs.xls 일위대가 파싱 완료")
        print(f"   통합 구조 JSON 저장: {output_file}")

if __name__ == "__main__":
    main()