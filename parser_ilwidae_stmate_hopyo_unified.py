"""
stmate.xlsx 일위대가_호표 시트 통합 파서
통합 JSON 구조: 품명, 규격, 단위, 수량 4개 컬럼으로 표준화
No.N 패턴 처리
"""
import pandas as pd
import re
import sys
import io
import xlwings as xw
from ilwidae_base_parser import IlwidaeBaseParser

# UTF-8 인코딩 설정
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except:
    pass  # 이미 설정되어 있거나 가상환경에서 실행 중

class IlwidaeStmateHopyoParser(IlwidaeBaseParser):
    """stmate.xlsx 일위대가_호표 파서"""
    
    def __init__(self):
        super().__init__('stmate.xlsx', '일위대가_호표')
        self.app = None
        self.wb = None
        self.ws = None
        
    def read_list_info(self):
        """일위대가목록 읽기 (No. 패턴이 아닌 항목들)"""
        try:
            if not self.app:
                self.app = xw.App(visible=False)
                self.wb = self.app.books.open(self.file_path)
            
            ws = self.wb.sheets['일위대가목록']
            
            hopyo_info = []
            used_range = ws.used_range
            
            if used_range:
                for row_idx in range(3, used_range.last_cell.row + 1):  # 3행부터 시작
                    try:
                        cell_value = ws.range(row_idx, 1).value
                        if cell_value:
                            cell_str = str(cell_value).strip()
                            # #N 형식이 아닌 다른 패턴 (일위대가_호표용)
                            if not cell_str.startswith('#'):
                                # 일위대가 항목으로 추정
                                spec = ws.range(row_idx, 2).value or ''
                                unit = ws.range(row_idx, 4).value or ''
                                
                                hopyo_info.append({
                                    'num': str(len(hopyo_info) + 1),  # 순서대로 번호 부여
                                    'work': cell_str,
                                    'spec': str(spec).strip() if spec else '',
                                    'unit': str(unit).strip() if unit else ''
                                })
                    except:
                        continue
            
            print(f"\n[일위대가목록 정보]")
            print(f"총 {len(hopyo_info)}개 일위대가 항목 발견")
            
            return hopyo_info
            
        except Exception as e:
            print(f"목록 읽기 오류: {str(e)}")
            return None
    
    def find_ilwidae_hopos(self):
        """일위대가 호표 찾기 (No.N 패턴)"""
        hopyo_list = []
        
        try:
            # xlwings로 엑셀 파일 읽기
            if not self.app:
                self.app = xw.App(visible=False)
                self.wb = self.app.books.open(self.file_path)
            
            self.ws = self.wb.sheets[self.sheet_name]
            used_range = self.ws.used_range
            
            if used_range:
                print(f"\n시트 '{self.sheet_name}' 로드 완료. 크기: {used_range.last_cell.row} x {used_range.last_cell.column}")
                
                # DataFrame으로 변환
                data = []
                for row_idx in range(1, min(used_range.last_cell.row + 1, 1000)):
                    row_data = []
                    for col_idx in range(1, min(used_range.last_cell.column + 1, 20)):
                        cell_value = self.ws.range(row_idx, col_idx).value
                        row_data.append(cell_value)
                    data.append(row_data)
                
                self.df = pd.DataFrame(data)
                
                # 컬럼 감지
                self.detect_columns(self.df)
                print(f"컬럼 매핑: {self.column_map}")
            else:
                print(f"시트 '{self.sheet_name}'에 데이터가 없습니다.")
                return []
            
            # 호표 패턴: No.N 작업명
            hopyo_pattern = r'^No\.(\d+)\s+(.+)'
            
            # 호표 찾기
            for row_idx in range(1, used_range.last_cell.row + 1):
                try:
                    cell_value = self.ws.range(row_idx, 1).value
                    if cell_value:
                        cell_str = str(cell_value).strip()
                        match = re.match(hopyo_pattern, cell_str)
                        if match:
                            hopyo_num = match.group(1)
                            work_name = match.group(2).strip()
                            
                            hopyo_list.append({
                                'row': row_idx - 1,  # DataFrame 인덱스로 변환
                                'num': hopyo_num,
                                'work': work_name,
                                'full_text': cell_str
                            })
                except:
                    continue
            
        except ImportError:
            print("xlwings가 설치되지 않았습니다. stmate 파일은 xlwings가 필요합니다.")
            print("설치: pip install xlwings")
            return []
        except Exception as e:
            print(f"파일 읽기 오류: {str(e)}")
            return []
        
        return hopyo_list
    
    def parse_ilwidae_title(self, row_idx, df):
        """일위대가 타이틀 파싱 (No. 패턴 특화)"""
        title_data = {
            "품명": "",
            "규격": "",
            "단위": "",
            "수량": ""
        }
        
        # 호표 행에서 작업명 추출
        cell = df.iloc[row_idx, 0]
        if pd.notna(cell):
            # No.N 패턴 제거하고 작업명만 추출
            work_name = re.sub(r'^No\.\d+\s+', '', str(cell).strip())
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
    
    def __del__(self):
        """소멸자: xlwings 리소스 정리"""
        try:
            if self.wb:
                self.wb.close()
            if self.app:
                self.app.quit()
        except:
            pass

def main():
    """메인 실행 함수"""
    parser = IlwidaeStmateHopyoParser()
    
    # 파싱 및 저장
    output_file = parser.save_to_json()
    
    if output_file:
        print(f"\n✅ stmate.xlsx 일위대가_호표 파싱 완료")
        print(f"   통합 구조 JSON 저장: {output_file}")

if __name__ == "__main__":
    main()