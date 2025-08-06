import pandas as pd
import json
import sys
import io
import re
import os
import subprocess

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def detect_file_type(file_path):
    """파일 구조 자동 감지"""
    
    print(f"\n파일 구조 감지 중: {file_path}")
    
    # 엑셀 파일 열기
    if file_path.endswith('.xls'):
        xls = pd.ExcelFile(file_path, engine='xlrd')
    else:
        xls = pd.ExcelFile(file_path)
    
    # 일위대가 시트 찾기
    target_sheet = None
    for sheet in xls.sheet_names:
        if '일위대가' in sheet and '목록' not in sheet and '총괄' not in sheet:
            target_sheet = sheet
            break
    
    if not target_sheet:
        for sheet in xls.sheet_names:
            if '산근' in sheet:
                target_sheet = sheet
                break
    
    if not target_sheet:
        return None, None
    
    print(f"대상 시트: {target_sheet}")
    
    # 시트 읽기
    if file_path.endswith('.xls'):
        df = pd.read_excel(file_path, sheet_name=target_sheet, header=None, engine='xlrd')
    else:
        df = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
    
    # 호표 패턴으로 타입 결정
    for i in range(min(30, len(df))):
        for j in range(min(5, len(df.columns))):
            cell = df.iloc[i, j]
            if pd.notna(cell) and '호표' in str(cell):
                cell_str = str(cell)
                
                # 패턴별 분류
                if '：' in cell_str or ':' in cell_str:
                    print("감지된 타입: SGS 형식 (콜론 구분)")
                    return 'sgs', target_sheet
                elif '(' in cell_str and ')' in cell_str and '호표' in cell_str:
                    print("감지된 타입: 건축구조 형식 (괄호 구분)")
                    return 'construction', target_sheet
                elif j == 1 and '제' in cell_str:
                    print("감지된 타입: EST 형식 (1열 호표)")
                    return 'est', target_sheet
                elif j == 0 and '제' in cell_str:
                    print("감지된 타입: TEST1 형식 (0열 호표)")
                    return 'test1', target_sheet
    
    return None, None

def parse_with_appropriate_parser(file_path):
    """적절한 파서로 파일 파싱"""
    
    file_type, sheet_name = detect_file_type(file_path)
    
    if not file_type:
        print("파일 타입을 감지할 수 없습니다.")
        return None
    
    print(f"\n{file_type} 파서 실행 중...")
    
    # 파서 매핑
    parser_map = {
        'test1': 'parser_test1.py',
        'sgs': 'parser_sgs.py',
        'construction': 'parser_construction.py',
        'est': 'parser_est.py'
    }
    
    parser_file = parser_map.get(file_type)
    if not parser_file:
        print(f"파서를 찾을 수 없습니다: {file_type}")
        return None
    
    # 원본 파일 백업 (필요한 경우)
    backup_needed = False
    backup_file = None
    
    # 각 파서가 기대하는 파일명
    expected_names = {
        'test1': 'test1.xlsx',
        'sgs': 'sgs.xls',
        'construction': '건축구조내역2.xlsx',
        'est': 'est.xlsx'
    }
    
    expected_name = expected_names.get(file_type)
    
    # 파일명이 다르면 임시로 복사
    if file_path != expected_name:
        import shutil
        
        # 기존 파일이 있으면 백업
        if os.path.exists(expected_name):
            backup_file = expected_name + '.backup'
            shutil.move(expected_name, backup_file)
            backup_needed = True
        
        # 복사
        shutil.copy(file_path, expected_name)
        print(f"임시 파일 생성: {expected_name}")
    
    # 파서 실행
    try:
        result = subprocess.run(['python', parser_file], 
                              capture_output=True, 
                              text=True, 
                              encoding='utf-8',
                              errors='replace')
        
        if result.returncode == 0:
            print("파싱 성공!")
            
            # 생성된 JSON 파일 읽기
            json_files = {
                'test1': 'test1_parsed.json',
                'sgs': 'sgs_parsed.json',
                'construction': 'construction_parsed.json',
                'est': 'est_parsed.json'
            }
            
            json_file = json_files.get(file_type)
            if json_file and os.path.exists(json_file):
                with open(json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # 감지된 타입 추가
                data['detected_type'] = file_type
                data['original_file'] = file_path
                
                # 새 파일로 저장
                output_file = os.path.splitext(file_path)[0] + '_smart.json'
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                
                print(f"\n결과 저장: {output_file}")
                print(f"호표 수: {data.get('hopyo_count', 0)}")
                
                total_items = 0
                if 'hopyo_data' in data:
                    total_items = sum(d.get('항목수', 0) for d in data['hopyo_data'].values())
                print(f"총 세부항목: {total_items}")
                
                return data
            else:
                print(f"JSON 파일을 찾을 수 없습니다: {json_file}")
        else:
            print(f"파서 실행 실패: {result.stderr}")
    
    finally:
        # 임시 파일 정리
        if file_path != expected_name and os.path.exists(expected_name):
            os.remove(expected_name)
            print(f"임시 파일 삭제: {expected_name}")
        
        # 백업 복원
        if backup_needed and backup_file and os.path.exists(backup_file):
            shutil.move(backup_file, expected_name)
    
    return None

def main():
    """메인 함수 - 사용 예시"""
    
    print("="*80)
    print("스마트 파서 - 파일 구조 자동 감지 및 파싱")
    print("="*80)
    
    # 예시: 업로드된 파일 처리
    if len(sys.argv) > 1:
        # 명령줄 인자로 파일 경로 받기
        file_path = sys.argv[1]
        if os.path.exists(file_path):
            parse_with_appropriate_parser(file_path)
        else:
            print(f"파일을 찾을 수 없습니다: {file_path}")
    else:
        # 테스트: 모든 파일 처리
        test_files = ['test1.xlsx', 'sgs.xls', '건축구조내역2.xlsx', 'est.xlsx']
        
        for file in test_files:
            if os.path.exists(file):
                print(f"\n\n{'='*60}")
                print(f"파일: {file}")
                print(f"{'='*60}")
                parse_with_appropriate_parser(file)

if __name__ == "__main__":
    main()