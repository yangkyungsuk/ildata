import pandas as pd
import os
import sys
import io
import re

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def find_header_info(df):
    """헤더 행 찾기 및 열 정보 반환"""
    header_keywords = ['산출근거', '산 출 근 거', '산출내역', '산 출 내 역']
    
    for i in range(min(10, len(df))):
        for j in range(len(df.columns)):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                for keyword in header_keywords:
                    if keyword in cell_str:
                        # 같은 행에서 다른 헤더들 찾기
                        headers = {}
                        for k in range(len(df.columns)):
                            h = df.iloc[i, k]
                            if pd.notna(h):
                                h_str = str(h).strip()
                                if h_str:
                                    headers[k] = h_str
                        return i, headers
    return None, {}

def analyze_sangcul_detail(file_path, sheet_name):
    """산출근거/산출내역 상세 구조 분석"""
    print(f"\n{'='*80}")
    print(f"파일: {file_path} | 시트: {sheet_name}")
    print('='*80)
    
    try:
        # 파일 읽기
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        
        # 헤더 찾기
        header_row, headers = find_header_info(df)
        
        if header_row is not None:
            print(f"\n헤더 행: {header_row}")
            print("헤더 내용:")
            for col, header in headers.items():
                print(f"  열 {col}: {header}")
            
            # 비고 열 찾기
            bigo_col = None
            for col, header in headers.items():
                if '비고' in header or '비 고' in header:
                    bigo_col = col
                    print(f"\n비고 열 발견: 열 {bigo_col}")
                    break
        
        # 호표 패턴 분석
        patterns = [
            (r'^(\d+)\.\s*(.+)', 'dot_style'),  # 1.작업명
            (r'산근\s*(\d+)\s*호표\s*[：:](.+)', 'sanggun_style'),  # 산근 1 호표：작업명
            (r'^#(\d+)\s+(.+)', 'hash_style'),  # #1 작업명
        ]
        
        # 첫 번째 호표 찾기
        first_hopyo = None
        for i in range(header_row + 1 if header_row else 0, min(100, len(df))):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    for pattern, style in patterns:
                        match = re.search(pattern, cell_str)
                        if match:
                            hopyo_num = match.group(1)
                            work_name = match.group(2) if len(match.groups()) > 1 else ''
                            
                            # 대분류인지 확인 (호표 번호가 1부터 순차적으로 증가)
                            if hopyo_num in ['1', '2', '3', '4', '5']:
                                first_hopyo = {
                                    'row': i, 'col': j, 'num': hopyo_num,
                                    'work': work_name, 'style': style,
                                    'text': cell_str
                                }
                                break
                if first_hopyo:
                    break
            if first_hopyo:
                break
        
        if first_hopyo:
            print(f"\n첫 번째 호표:")
            print(f"  스타일: {first_hopyo['style']}")
            print(f"  위치: 행 {first_hopyo['row']}, 열 {first_hopyo['col']}")
            print(f"  내용: {first_hopyo['text']}")
            
            # 호표 데이터 샘플
            print(f"\n호표 데이터 샘플 (행 {first_hopyo['row']} ~ {first_hopyo['row']+20}):")
            
            # 산출근거 열 찾기
            sangcul_col = 0  # 보통 0열
            for col, header in headers.items():
                if '산출' in header and ('근거' in header or '내역' in header):
                    sangcul_col = col
                    break
            
            for i in range(first_hopyo['row'], min(first_hopyo['row']+20, len(df))):
                # 산출근거 내용
                sangcul_data = df.iloc[i, sangcul_col] if sangcul_col < len(df.columns) else None
                bigo_data = df.iloc[i, bigo_col] if bigo_col and bigo_col < len(df.columns) else None
                
                row_info = []
                if pd.notna(sangcul_data):
                    row_info.append(f"산출근거: {str(sangcul_data).strip()}")
                if pd.notna(bigo_data):
                    row_info.append(f"비고: {str(bigo_data).strip()}")
                
                if row_info:
                    print(f"  행 {i}: {' | '.join(row_info)}")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    # 분석할 파일들
    files_to_analyze = [
        ('test1.xlsx', '단가산출_산근'),
        ('est.xlsx', '단가산출'),
        ('sgs.xls', '단가산출'),
        ('ebs.xls', '일위대가_산근'),
        ('건축구조내역.xlsx', '중기단가산출'),  # 산출내역으로 되어있을 가능성
        ('건축구조내역2.xlsx', '중기단가산출'),  # 확인용
    ]
    
    print("산출근거/산출내역 구조 분석")
    print("="*80)
    
    for file_name, sheet_name in files_to_analyze:
        if os.path.exists(file_name):
            analyze_sangcul_detail(file_name, sheet_name)
        else:
            print(f"\n파일을 찾을 수 없음: {file_name}")

if __name__ == "__main__":
    main()