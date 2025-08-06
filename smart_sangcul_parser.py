"""
스마트 단가산출 파서
파일을 자동 인식하여 적절한 파서 실행
"""
import os
import sys
import subprocess
import pandas as pd

def detect_file_type(file_path):
    """파일 타입과 적절한 파서 결정"""
    file_name = os.path.basename(file_path).lower()
    
    # 파일명으로 우선 판단 (v4 파서 우선 사용)
    if 'test1' in file_name:
        return 'parser_sangcul_test1_v4.py'
    elif 'sgs' in file_name:
        return 'parser_sangcul_sgs_v4.py'
    elif 'stmate' in file_name:
        return 'parser_sangcul_stmate_v4.py'
    elif 'est' in file_name and 'stmate' not in file_name:
        return 'parser_sangcul_est_v4.py'
    elif 'ebs' in file_name:
        return 'parser_sangcul_ebs_v4.py'
    elif '건축구조내역2' in file_name:
        return 'parser_sangcul_construction2_v4.py'
    elif '건축구조내역' in file_name:
        return 'parser_sangcul_construction_v4.py'
    
    # 파일명으로 판단 안되면 시트 구조로 판단
    try:
        if file_path.endswith('.xls'):
            xls = pd.ExcelFile(file_path, engine='xlrd')
        else:
            xls = pd.ExcelFile(file_path)
        
        sheets = xls.sheet_names
        
        # 시트 이름으로 판단 (v4 파서 우선 사용)
        if '단가산출_산근' in sheets:
            return 'parser_sangcul_test1_v4.py'
        elif '중기단가산출서' in sheets:
            return 'parser_sangcul_construction_v4.py'
        elif '단가산출서' in sheets:
            return 'parser_sangcul_construction2_v4.py'
        elif '일위대가_산근' in sheets and '일위대가목록' in sheets:
            # ebs 스타일 (목록에 #N 형식 확인)
            df = pd.read_excel(file_path, sheet_name='일위대가목록', nrows=10, header=None)
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    cell = df.iloc[i, j]
                    if pd.notna(cell) and str(cell).strip().startswith('#'):
                        return 'parser_sangcul_ebs_v4.py'
        elif '단가산출' in sheets:
            # 산근 스타일인지 확인
            df = pd.read_excel(file_path, sheet_name='단가산출', nrows=10, header=None)
            for i in range(len(df)):
                cell = df.iloc[i, 0]
                if pd.notna(cell) and '산근' in str(cell) and '호표' in str(cell):
                    return 'parser_sangcul_sgs_v4.py'
            return 'parser_sangcul_est_v4.py'
            
    except Exception as e:
        print(f"파일 타입 감지 오류: {str(e)}")
    
    return None

def run_parser(file_path):
    """적절한 파서 실행"""
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return False
    
    # 파일 타입 감지
    parser = detect_file_type(file_path)
    
    if not parser:
        print(f"적절한 파서를 찾을 수 없습니다: {file_path}")
        return False
    
    print(f"파일: {file_path}")
    print(f"감지된 파서: {parser}")
    print("-" * 50)
    
    # 해당 파서가 있는지 확인
    if not os.path.exists(parser):
        # v4가 없으면 v3 시도
        parser_v3 = parser.replace('_v4.py', '_v3.py')
        if os.path.exists(parser_v3):
            parser = parser_v3
        else:
            # v3도 없으면 기본 버전 시도
            parser_base = parser.replace('_v4.py', '.py')
            if os.path.exists(parser_base):
                parser = parser_base
            else:
                print(f"파서 파일을 찾을 수 없습니다: {parser}")
                return False
    
    # 파서 실행
    try:
        # 파일명을 인자로 전달할 수 있는 파서는 직접 전달
        # 기존 파서들은 파일명이 하드코딩되어 있으므로 해당 파일이 있어야 함
        result = subprocess.run([sys.executable, parser], 
                              capture_output=True, 
                              text=True)
        
        print(result.stdout)
        if result.stderr:
            print("오류:", result.stderr)
        
        return result.returncode == 0
        
    except Exception as e:
        print(f"파서 실행 오류: {str(e)}")
        return False

def main():
    """메인 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(description='스마트 단가산출 파서')
    parser.add_argument('file', nargs='?', help='파싱할 Excel 파일')
    parser.add_argument('--all', action='store_true', help='모든 파일 파싱')
    
    args = parser.parse_args()
    
    if args.all:
        # 모든 파일 파싱
        files = [
            'test1.xlsx',
            'sgs.xls', 
            'est.xlsx',
            'ebs.xls',
            'stmate.xlsx',
            '건축구조내역.xlsx',
            '건축구조내역2.xlsx'
        ]
        
        success_count = 0
        for file in files:
            if os.path.exists(file):
                print(f"\n{'='*60}")
                if run_parser(file):
                    success_count += 1
                print('='*60)
            else:
                print(f"\n파일 없음: {file}")
        
        print(f"\n완료: {success_count}개 파일 파싱 성공")
        
    elif args.file:
        # 특정 파일만 파싱
        run_parser(args.file)
    else:
        # 현재 폴더의 Excel 파일 목록 표시
        excel_files = [f for f in os.listdir('.') 
                      if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        
        if excel_files:
            print("현재 폴더의 Excel 파일:")
            for i, file in enumerate(excel_files, 1):
                print(f"{i}. {file}")
            
            print("\n사용법:")
            print(f"python {os.path.basename(__file__)} <파일명>")
            print(f"python {os.path.basename(__file__)} --all")
        else:
            print("현재 폴더에 Excel 파일이 없습니다.")

if __name__ == "__main__":
    main()