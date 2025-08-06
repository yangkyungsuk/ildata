"""
stmate.xlsx 파일 구조 분석 (xlwings 사용)
"""
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_stmate_with_xlwings():
    """xlwings로 stmate.xlsx 분석"""
    file_path = 'stmate.xlsx'
    
    try:
        import xlwings as xw
        
        # 백그라운드에서 Excel 실행
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        
        sheets = [sheet.name for sheet in wb.sheets]
        print(f"시트 목록: {sheets}")
        
        for sheet_name in sheets:
            print(f"\n{'='*50}")
            print(f"시트: {sheet_name}")
            print('='*50)
            
            try:
                ws = wb.sheets[sheet_name]
                used_range = ws.used_range
                if used_range:
                    print(f"사용된 범위: {used_range.address}")
                    
                    # 처음 20행, 10열 읽기
                    print("\n처음 20행 내용:")
                    for row_idx in range(1, min(21, used_range.last_cell.row + 1)):
                        row_content = []
                        for col_idx in range(1, min(11, used_range.last_cell.column + 1)):
                            try:
                                cell_value = ws.range(row_idx, col_idx).value
                                if cell_value is not None:
                                    content = str(cell_value).strip()
                                    if content:
                                        row_content.append(f"[{col_idx}]:{content[:30]}")
                            except:
                                continue
                        if row_content:
                            print(f"행{row_idx}: {' | '.join(row_content)}")
                
                # 호표 패턴 찾기
                print(f"\n호표 패턴 검색:")
                import re
                patterns = [
                    r'제\d+호표',      # 제N호표
                    r'호표\s*\d+',     # 호표 N
                    r'\d+\.\s*',       # N. 형식
                    r'산근\s*\d+',     # 산근 N
                    r'#\d+',           # #N
                    r'\(\s*산근\s*\d+\s*\)'  # (산근 N)
                ]
                
                found_patterns = []
                
                if used_range:
                    for row_idx in range(1, min(101, used_range.last_cell.row + 1)):
                        for col_idx in range(1, min(11, used_range.last_cell.column + 1)):
                            try:
                                cell_value = ws.range(row_idx, col_idx).value
                                if cell_value is not None:
                                    cell_str = str(cell_value).strip()
                                    for pattern in patterns:
                                        if re.search(pattern, cell_str):
                                            found_patterns.append({
                                                'row': row_idx,
                                                'col': col_idx,
                                                'pattern': pattern,
                                                'text': cell_str[:50]
                                            })
                            except:
                                continue
                
                if found_patterns:
                    print("발견된 호표 패턴:")
                    for p in found_patterns[:10]:
                        print(f"  행{p['row']}, 열{p['col']}: {p['pattern']} -> {p['text']}")
                else:
                    print("호표 패턴을 찾을 수 없습니다.")
                
            except Exception as e:
                print(f"시트 '{sheet_name}' 분석 오류: {e}")
        
        wb.close()
        app.quit()
    
    except ImportError:
        print("xlwings가 설치되지 않았습니다. pip install xlwings로 설치하세요.")
    except Exception as e:
        print(f"파일 분석 오류: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_stmate_with_xlwings()