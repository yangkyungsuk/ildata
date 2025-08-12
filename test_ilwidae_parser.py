"""
간단한 일위대가 파서 테스트
stmate.xlsx 파일을 pandas로 직접 읽어서 처리
"""
import pandas as pd
import json
import os

def test_parse_stmate():
    """stmate.xlsx 테스트 파싱"""
    
    # result 폴더 생성
    os.makedirs('result', exist_ok=True)
    
    file_path = 'stmate.xlsx'
    
    # 사용 가능한 시트 확인
    try:
        xl_file = pd.ExcelFile(file_path)
        print(f"파일: {file_path}")
        print(f"시트 목록: {xl_file.sheet_names}")
        print("-" * 50)
    except Exception as e:
        print(f"파일 읽기 오류: {str(e)}")
        return
    
    # 각 일위대가 관련 시트 처리
    ilwidae_sheets = []
    for sheet_name in xl_file.sheet_names:
        if '일위대가' in sheet_name:
            ilwidae_sheets.append(sheet_name)
    
    print(f"일위대가 관련 시트: {ilwidae_sheets}")
    print("-" * 50)
    
    # 각 시트 데이터 읽기 및 저장
    for sheet_name in ilwidae_sheets:
        try:
            print(f"\n[{sheet_name}] 처리 중...")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            print(f"  - 크기: {df.shape}")
            
            # 첫 10행 확인
            print(f"  - 첫 5행 내용:")
            for i in range(min(5, len(df))):
                row_str = ""
                for j in range(min(5, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        row_str += f"{str(cell)[:20]:20} | "
                if row_str:
                    print(f"    행{i}: {row_str}")
            
            # 간단한 JSON 구조로 저장
            result = {
                'file': file_path,
                'sheet': sheet_name,
                'rows': len(df),
                'columns': len(df.columns),
                'sample_data': []
            }
            
            # 샘플 데이터 추가
            for i in range(min(20, len(df))):
                row_data = []
                for j in range(min(10, len(df.columns))):
                    cell = df.iloc[i, j]
                    if pd.notna(cell):
                        row_data.append(str(cell))
                    else:
                        row_data.append("")
                
                result['sample_data'].append({
                    'row': i,
                    'data': row_data
                })
            
            # JSON 파일로 저장
            output_file = f"result/{file_path.split('.')[0]}_{sheet_name.replace(' ', '_')}_sample.json"
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            
            print(f"  ✓ 저장 완료: {output_file}")
            
        except Exception as e:
            print(f"  ✗ 처리 실패: {str(e)}")

def test_other_files():
    """다른 엑셀 파일들도 확인"""
    excel_files = ['test1 (2).xlsx', 'est.xlsx', 'sgs.xls', '건축구조내역.xlsx', '건축구조내역2.xlsx']
    
    print("\n" + "=" * 70)
    print("다른 엑셀 파일 확인")
    print("=" * 70)
    
    for file_path in excel_files:
        if os.path.exists(file_path):
            try:
                xl_file = pd.ExcelFile(file_path)
                print(f"\n파일: {file_path}")
                print(f"시트 목록: {xl_file.sheet_names}")
                
                # 일위대가 관련 시트 찾기
                ilwidae_sheets = [s for s in xl_file.sheet_names if '일위대가' in s]
                if ilwidae_sheets:
                    print(f"  → 일위대가 시트: {ilwidae_sheets}")
                else:
                    print(f"  → 일위대가 시트 없음")
                    
            except Exception as e:
                print(f"파일 읽기 실패: {file_path} - {str(e)}")
        else:
            print(f"파일 없음: {file_path}")

if __name__ == "__main__":
    print("=" * 70)
    print("일위대가 파서 테스트")
    print("=" * 70)
    print()
    
    # stmate.xlsx 테스트
    test_parse_stmate()
    
    # 다른 파일들 확인
    test_other_files()
    
    print("\n" + "=" * 70)
    print("테스트 완료")
    print("=" * 70)