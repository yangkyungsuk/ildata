import pandas as pd
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def check_headers():
    """모든 파일의 헤더 구조 확인"""
    
    files = [
        ('test1.xlsx', '일위대가_산근'),
        ('sgs.xls', '일위대가'),
        ('건축구조내역2.xlsx', '일위대가'),
        ('est.xlsx', '일위대가')
    ]
    
    print("="*80)
    print("모든 파일의 헤더 구조 확인")
    print("="*80)
    
    for file_path, sheet_name in files:
        print(f"\n\n{'='*60}")
        print(f"{file_path} - {sheet_name}")
        print(f"{'='*60}")
        
        try:
            # 엑셀 읽기
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            print(f"파일 크기: {df.shape[0]}행 x {df.shape[1]}열")
            print("\n처음 15행 (주요 열만):")
            print("-" * 60)
            
            # 처음 15행 출력
            for row_idx in range(min(15, len(df))):
                row_data = []
                
                # 처음 10열만 확인
                for col_idx in range(min(10, len(df.columns))):
                    val = df.iloc[row_idx, col_idx]
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if val_str:
                            # 긴 텍스트는 축약
                            if len(val_str) > 30:
                                val_str = val_str[:27] + "..."
                            row_data.append(f"[{col_idx}]{val_str}")
                
                if row_data:
                    print(f"행{row_idx:2d}: {' | '.join(row_data)}")
            
            # 호표 패턴 찾기
            print("\n호표 패턴:")
            print("-" * 40)
            hopyo_found = False
            
            for row_idx in range(min(30, len(df))):
                for col_idx in range(min(5, len(df.columns))):
                    val = df.iloc[row_idx, col_idx]
                    if pd.notna(val):
                        val_str = str(val)
                        if '호표' in val_str and ('제' in val_str or '(' in val_str):
                            print(f"  행{row_idx}, 열{col_idx}: {val_str[:50]}")
                            hopyo_found = True
                            break
                if hopyo_found:
                    break
            
        except Exception as e:
            print(f"오류 발생: {str(e)}")
    
    print("\n" + "="*80)
    print("헤더 구조 확인 완료")
    print("="*80)

if __name__ == "__main__":
    check_headers()