"""
stmate.xlsx 일위대가_호표 시트 파서 v4 (완전한 순서 보존)
일위대가 구조: No.1, No.2, ... 형식
규칙: 호표 사이의 모든 행을 순서대로 빠짐없이 JSON에 저장
"""
import pandas as pd
import json
import re
import sys
import io
import xlwings as xw

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def read_ilwidae_list(file_path):
    """일위대가목록에서 호표 정보 읽기 (xlwings 사용)"""
    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        ws = wb.sheets['일위대가목록']
        
        hopyo_info = []
        used_range = ws.used_range
        
        if used_range:
            for row_idx in range(3, used_range.last_cell.row + 1):  # 3행부터 시작 (헤더 제외)
                try:
                    cell_value = ws.range(row_idx, 1).value
                    if cell_value:
                        cell_str = str(cell_value).strip()
                        # #N 형식이 아닌 다른 패턴도 체크 (일위대가용)
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
        
        wb.close()
        app.quit()
        
        print(f"\n[일위대가목록 정보]")
        print(f"총 {len(hopyo_info)}개 일위대가 항목 발견")
        
        return hopyo_info
        
    except Exception as e:
        print(f"목록 읽기 오류: {str(e)}")
        return None

def extract_all_rows_in_order_xlwings(ws, start_row, end_row):
    """xlwings로 호표 사이의 모든 행을 순서대로 완전히 추출"""
    all_rows = []
    
    for row_idx in range(start_row, end_row):
        # 해당 행의 모든 컬럼 데이터 수집
        row_data = {
            'row_number': row_idx,
            'columns': {},
            'has_content': False
        }
        
        # 처음 20컬럼 검사
        for col_idx in range(1, 21):
            try:
                cell_value = ws.range(row_idx, col_idx).value
                if cell_value is not None:
                    content = str(cell_value).strip()
                    if content:  # 빈 문자열이 아니면 저장
                        row_data['columns'][f'col_{col_idx-1}'] = content
                        row_data['has_content'] = True
            except:
                continue
        
        # 빈 행이어도 순서 유지를 위해 모두 저장
        all_rows.append(row_data)
    
    return all_rows

def parse_stmate_ilwidae():
    """stmate.xlsx의 일위대가_호표 시트 파싱 (완전한 순서 보존)"""
    
    file_path = 'stmate.xlsx'
    sheet_name = '일위대가_호표'
    
    try:
        # 목록 정보 읽기
        list_info = read_ilwidae_list(file_path)
        
        # xlwings로 엑셀 파일 읽기
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        ws = wb.sheets[sheet_name]
        
        used_range = ws.used_range
        if used_range:
            print(f"\n시트 '{sheet_name}' 로드 완료. 크기: {used_range.last_cell.row} x {used_range.last_cell.column}")
        else:
            print(f"\n시트 '{sheet_name}'에 데이터가 없습니다.")
            wb.close()
            app.quit()
            return None
        
        # 호표 패턴: No.N 작업명
        hopyo_pattern = r'^No\.(\d+)\s+(.+)'
        
        # 호표 찾기
        hopyo_list = []
        for row_idx in range(1, used_range.last_cell.row + 1):
            try:
                cell_value = ws.range(row_idx, 1).value
                if cell_value:
                    cell_str = str(cell_value).strip()
                    match = re.match(hopyo_pattern, cell_str)
                    if match:
                        hopyo_num = match.group(1)
                        work_name = match.group(2).strip()
                        
                        hopyo_list.append({
                            'row': row_idx,
                            'num': hopyo_num,
                            'work': work_name,
                            'full_text': cell_str
                        })
            except:
                continue
        
        print(f"\n[파싱 결과]")
        print(f"총 {len(hopyo_list)}개 호표 발견")
        
        # 목록과 비교
        if list_info:
            print(f"\n[검증 결과]")
            print(f"목록표: {len(list_info)}개, 실제 파싱: {len(hopyo_list)}개")
            print("✓ 호표 개수 일치" if len(list_info) == len(hopyo_list) else "✗ 호표 개수 불일치")
        
        # 결과 구성
        result = {
            'file': file_path,
            'sheet': sheet_name,
            'total_hopyo_count': len(hopyo_list),
            'validation': {
                'list_count': len(list_info) if list_info else 0,
                'parsed_count': len(hopyo_list),
                'match': len(list_info) == len(hopyo_list) if list_info else False
            },
            'hopyo_data': {}
        }
        
        # 각 호표의 완전한 데이터 추출
        for idx, hopyo in enumerate(hopyo_list):
            hopyo_key = f"호표{hopyo['num']}"
            
            # 호표 범위 설정
            start_row = hopyo['row']
            end_row = hopyo_list[idx + 1]['row'] if idx + 1 < len(hopyo_list) else used_range.last_cell.row + 1
            
            # 호표 사이의 모든 행을 순서대로 완전히 추출
            all_rows_data = extract_all_rows_in_order_xlwings(ws, start_row, end_row)
            
            # 내용이 있는 행만 계산 (통계용)
            content_rows = [row for row in all_rows_data if row['has_content']]
            
            result['hopyo_data'][hopyo_key] = {
                '호표번호': hopyo['num'],
                '작업명': hopyo['work'],
                '시작행': start_row,
                '종료행': end_row,
                '총_행수': len(all_rows_data),
                '내용있는_행수': len(content_rows),
                '모든_행_데이터': all_rows_data  # 순서대로 모든 행 저장
            }
            
            print(f"\n호표{hopyo['num']}: {hopyo['work']}")
            print(f"  - 전체 행 범위: {start_row} ~ {end_row-1} ({len(all_rows_data)}개 행)")
            print(f"  - 내용이 있는 행: {len(content_rows)}개")
            
            # 처음 5개 행 샘플 출력
            print(f"  - 처음 5개 행 샘플:")
            for i, row_data in enumerate(all_rows_data[:5]):
                if row_data['has_content']:
                    main_content = list(row_data['columns'].values())[0] if row_data['columns'] else ''
                    print(f"    행{row_data['row_number']}: \"{main_content[:40]}...\"")
                else:
                    print(f"    행{row_data['row_number']}: (빈행)")
        
        wb.close()
        app.quit()
        
        # JSON 파일로 저장
        output_file = 'stmate_ilwidae_parsed_v4.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n파싱 완료. 결과 저장: {output_file}")
        print(f"✅ 호표 사이의 모든 행이 순서대로 완전히 저장되었습니다.")
        
        return result
        
    except ImportError:
        print("xlwings가 설치되지 않았습니다. stmate 파일은 xlwings가 필요합니다.")
        return None
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    parse_stmate_ilwidae()