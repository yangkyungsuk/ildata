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

class CompleteAutoParser:
    """완전한 자동 파서 - 모든 타입 구현"""
    
    def detect_file_type(self, file_path: str) -> Tuple[str, str]:
        """파일 구조를 자동으로 감지"""
        
        print(f"\n파일 구조 감지 중...")
        
        # 시트 목록 확인
        if file_path.endswith('.xls'):
            xls = pd.ExcelFile(file_path, engine='xlrd')
        else:
            xls = pd.ExcelFile(file_path)
        
        # 일위대가 시트 찾기 (목록/총괄 제외)
        target_sheet = None
        for sheet in xls.sheet_names:
            if '일위대가' in sheet and '목록' not in sheet and '총괄' not in sheet:
                target_sheet = sheet
                break
        
        # 못 찾으면 산근 시트
        if not target_sheet:
            for sheet in xls.sheet_names:
                if '산근' in sheet:
                    target_sheet = sheet
                    break
        
        if not target_sheet:
            return None, None
        
        # 시트 읽기
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=None, engine='xlrd')
        else:
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
        
        # 호표 패턴 확인
        for i in range(min(30, len(df))):
            for j in range(min(5, len(df.columns))):
                cell = df.iloc[i, j]
                if pd.notna(cell) and '호표' in str(cell):
                    cell_str = str(cell)
                    
                    # 패턴 분류
                    if '：' in cell_str or ':' in cell_str:
                        return 'sgs', target_sheet
                    elif '(' in cell_str and ')' in cell_str:
                        return 'construction', target_sheet
                    elif j == 1 and '제' in cell_str:
                        return 'est', target_sheet
                    elif j == 0 and '제' in cell_str:
                        return 'test1', target_sheet
        
        return None, None
    
    def parse_file(self, file_path: str) -> Dict[str, Any]:
        """파일을 자동으로 파싱"""
        
        file_type, sheet_name = self.detect_file_type(file_path)
        
        if not file_type:
            return {"error": "파일 구조를 인식할 수 없습니다"}
        
        print(f"감지된 타입: {file_type}")
        print(f"대상 시트: {sheet_name}")
        
        # 각 타입별 파싱
        if file_type == 'test1':
            return self.parse_test1(file_path, sheet_name)
        elif file_type == 'sgs':
            return self.parse_sgs(file_path, sheet_name)
        elif file_type == 'construction':
            return self.parse_construction(file_path, sheet_name)
        elif file_type == 'est':
            return self.parse_est(file_path, sheet_name)
    
    def parse_test1(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """test1 타입 파싱"""
        from parser_test1 import parse_test1
        
        # 파일이 이미 test1.xlsx가 아닌 경우만 복사
        import shutil
        if file_path != 'test1.xlsx':
            shutil.copy(file_path, 'test1.xlsx')
        parse_test1()
        
        # 생성된 JSON 읽기
        with open('test1_parsed.json', 'r', encoding='utf-8') as f:
            result = json.load(f)
        
        result['detected_type'] = 'test1'
        return result
    
    def parse_sgs(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """sgs 타입 파싱"""
        from parser_sgs import parse_sgs
        
        import shutil
        if file_path != 'sgs.xls':
            shutil.copy(file_path, 'sgs.xls')
        parse_sgs()
        
        with open('sgs_parsed.json', 'r', encoding='utf-8') as f:
            result = json.load(f)
        
        result['detected_type'] = 'sgs'
        return result
    
    def parse_construction(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """construction 타입 파싱"""
        from parser_construction import parse_construction
        
        import shutil
        if file_path != '건축구조내역2.xlsx':
            shutil.copy(file_path, '건축구조내역2.xlsx')
        parse_construction()
        
        with open('construction_parsed.json', 'r', encoding='utf-8') as f:
            result = json.load(f)
        
        result['detected_type'] = 'construction'
        return result
    
    def parse_est(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """est 타입 파싱"""
        from parser_est import parse_est
        
        import shutil
        if file_path != 'est.xlsx':
            shutil.copy(file_path, 'est.xlsx')
        parse_est()
        
        with open('est_parsed.json', 'r', encoding='utf-8') as f:
            result = json.load(f)
        
        result['detected_type'] = 'est'
        return result

def analyze_uploaded_file(file_path: str):
    """업로드된 파일 분석"""
    
    print("="*80)
    print(f"파일 분석: {file_path}")
    print("="*80)
    
    parser = CompleteAutoParser()
    result = parser.parse_file(file_path)
    
    if "error" not in result:
        # 결과 저장
        output_file = f"{os.path.splitext(file_path)[0]}_analyzed.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n분석 완료!")
        print(f"파일 타입: {result.get('detected_type', '알 수 없음')}")
        print(f"시트: {result.get('sheet', '알 수 없음')}")
        print(f"호표 수: {result.get('hopyo_count', 0)}")
        
        total_items = 0
        if 'hopyo_data' in result:
            total_items = sum(data.get('항목수', 0) for data in result['hopyo_data'].values())
        print(f"총 세부항목: {total_items}")
        print(f"\n결과 저장: {output_file}")
        
        return result
    else:
        print(f"\n오류: {result['error']}")
        return None

def main():
    """테스트 실행"""
    test_files = [
        'test1.xlsx',
        'sgs.xls',
        '건축구조내역2.xlsx',
        'est.xlsx'
    ]
    
    for file in test_files:
        if os.path.exists(file):
            print(f"\n\n{'='*60}")
            analyze_uploaded_file(file)
            print("="*60)

if __name__ == "__main__":
    main()