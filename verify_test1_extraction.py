"""
test1.xlsx 일위대가 추출 검증 스크립트
엑셀 파일과 생성된 JSON을 비교하여 정확성 검증
"""
import pandas as pd
import json
import re

def read_excel_data():
    """엑셀 파일 직접 읽기"""
    print("=" * 70)
    print("1. 엑셀 파일 원본 데이터 읽기")
    print("=" * 70)
    
    # 일위대가_산근 시트 읽기
    df = pd.read_excel('test1.xlsx', sheet_name='일위대가_산근', header=None)
    print(f"시트 크기: {df.shape[0]}행 x {df.shape[1]}열")
    
    # 일위대가목록표 읽기
    df_list = pd.read_excel('test1.xlsx', sheet_name='일위대가목록표', header=None)
    print(f"목록표 크기: {df_list.shape[0]}행 x {df_list.shape[1]}열")
    
    # 모든 호표 찾기
    print("\n[엑셀에서 직접 찾은 호표]")
    hopyo_pattern = r'제\s*(\d+)\s*호표'
    excel_hopyos = []
    
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = df.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(hopyo_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    
                    # 같은 행에서 작업명 찾기
                    work_name = ''
                    for k in range(j+1, len(df.columns)):
                        next_cell = df.iloc[i, k]
                        if pd.notna(next_cell):
                            next_str = str(next_cell).strip()
                            if next_str and not next_str.replace('.', '').replace(',', '').isdigit():
                                work_name = next_str
                                break
                    
                    excel_hopyos.append({
                        'num': hopyo_num,
                        'row': i,
                        'work': work_name,
                        'full_row_data': []
                    })
                    
                    # 전체 행 데이터 수집
                    for col in range(len(df.columns)):
                        cell_val = df.iloc[i, col]
                        if pd.notna(cell_val):
                            excel_hopyos[-1]['full_row_data'].append(str(cell_val).strip())
                    
                    print(f"  제{hopyo_num}호표 (행 {i+1}): {work_name}")
                    break
    
    # 각 호표의 상세 데이터 확인
    print("\n[각 호표의 산출근거 데이터 확인]")
    for idx, hopyo in enumerate(excel_hopyos):
        start_row = hopyo['row']
        end_row = excel_hopyos[idx + 1]['row'] if idx + 1 < len(excel_hopyos) else len(df)
        
        print(f"\n제{hopyo['num']}호표: {hopyo['work']} (행 {start_row+1} ~ {end_row})")
        print("  산출근거 항목:")
        
        # 산출근거 데이터 수집 (최대 10행)
        item_count = 0
        for row_idx in range(start_row + 1, min(start_row + 11, end_row)):
            row_data = []
            has_content = False
            
            for col_idx in range(min(10, len(df.columns))):
                cell = df.iloc[row_idx, col_idx]
                if pd.notna(cell):
                    value = str(cell).strip()
                    if value:
                        row_data.append(value)
                        has_content = True
                else:
                    row_data.append('')
            
            if has_content:
                item_count += 1
                # 주요 데이터 출력
                main_item = row_data[0] if row_data[0] else row_data[1] if len(row_data) > 1 else ''
                print(f"    행 {row_idx+1}: {main_item[:30]}")
                
                # 전체 행 내용 (디버깅용)
                if item_count <= 3:  # 처음 3개만 상세 출력
                    print(f"      상세: {' | '.join([v[:15] for v in row_data if v])}")
        
        print(f"  → 총 {item_count}개 항목")
    
    return df, excel_hopyos

def read_json_data():
    """생성된 JSON 파일 읽기"""
    print("\n" + "=" * 70)
    print("2. 생성된 JSON 파일 데이터 읽기")
    print("=" * 70)
    
    with open('result/test1_일위대가_산근_unified.json', 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    
    print(f"JSON 일위대가 개수: {json_data['total_ilwidae_count']}")
    
    json_hopyos = []
    for item in json_data['ilwidae_data']:
        json_hopyos.append({
            'num': item['ilwidae_no'],
            'work': item['ilwidae_title']['품명'],
            'start_row': item['position']['start_row'],
            'end_row': item['position']['end_row'],
            'sangul_count': len(item['산출근거'])
        })
        
        print(f"  제{item['ilwidae_no']}호표: {item['ilwidae_title']['품명']}")
        print(f"    - 위치: 행 {item['position']['start_row']+1} ~ {item['position']['end_row']}")
        print(f"    - 산출근거 항목 수: {len(item['산출근거'])}")
    
    return json_data, json_hopyos

def compare_data(excel_hopyos, json_hopyos, df, json_data):
    """엑셀과 JSON 데이터 비교"""
    print("\n" + "=" * 70)
    print("3. 데이터 비교 검증")
    print("=" * 70)
    
    # 호표 개수 비교
    print(f"\n[호표 개수 비교]")
    print(f"  엑셀: {len(excel_hopyos)}개")
    print(f"  JSON: {len(json_hopyos)}개")
    
    if len(excel_hopyos) == len(json_hopyos):
        print("  ✅ 호표 개수 일치")
    else:
        print("  ❌ 호표 개수 불일치!")
    
    # 각 호표별 상세 비교
    print(f"\n[호표별 상세 비교]")
    
    for i in range(min(len(excel_hopyos), len(json_hopyos))):
        excel_hopyo = excel_hopyos[i]
        json_hopyo = json_hopyos[i]
        
        print(f"\n제{excel_hopyo['num']}호표:")
        
        # 번호 일치 확인
        if excel_hopyo['num'] == json_hopyo['num']:
            print(f"  ✅ 번호 일치: {excel_hopyo['num']}")
        else:
            print(f"  ❌ 번호 불일치: 엑셀({excel_hopyo['num']}) vs JSON({json_hopyo['num']})")
        
        # 작업명 일치 확인
        if excel_hopyo['work'] == json_hopyo['work']:
            print(f"  ✅ 작업명 일치: {excel_hopyo['work']}")
        else:
            print(f"  ⚠️  작업명 차이:")
            print(f"     엑셀: {excel_hopyo['work']}")
            print(f"     JSON: {json_hopyo['work']}")
        
        # 위치 확인
        if excel_hopyo['row'] == json_hopyo['start_row']:
            print(f"  ✅ 시작 위치 일치: 행 {excel_hopyo['row']+1}")
        else:
            print(f"  ❌ 시작 위치 불일치: 엑셀(행 {excel_hopyo['row']+1}) vs JSON(행 {json_hopyo['start_row']+1})")
    
    # 산출근거 데이터 샘플 검증
    print(f"\n[산출근거 데이터 샘플 검증]")
    
    # 첫 번째 호표의 산출근거 상세 검증
    if len(json_data['ilwidae_data']) > 0:
        first_item = json_data['ilwidae_data'][0]
        print(f"\n제{first_item['ilwidae_no']}호표 산출근거 상세:")
        
        for sangul in first_item['산출근거'][:5]:  # 처음 5개만
            row_num = sangul['row_number']
            print(f"\n  행 {row_num+1}:")
            print(f"    품명: {sangul['품명']}")
            print(f"    수량: {sangul['수량']}")
            
            # 엑셀에서 같은 행 데이터 확인
            excel_row_data = []
            for col in range(min(10, len(df.columns))):
                cell = df.iloc[row_num, col]
                if pd.notna(cell):
                    excel_row_data.append(str(cell).strip())
            
            print(f"    엑셀 원본: {' | '.join(excel_row_data[:5])}")
            
            # 품명이 엑셀 데이터에 있는지 확인
            if sangul['품명'] in excel_row_data:
                print(f"    ✅ 품명이 엑셀 데이터에 존재")
            else:
                print(f"    ⚠️  품명이 엑셀 데이터와 다를 수 있음")

def check_data_completeness(df, json_data):
    """데이터 완전성 검사"""
    print("\n" + "=" * 70)
    print("4. 데이터 완전성 검사")
    print("=" * 70)
    
    # 빈 필드 확인
    print("\n[빈 필드 통계]")
    
    empty_counts = {
        '품명': 0,
        '규격': 0,
        '단위': 0,
        '수량': 0
    }
    
    total_sangul = 0
    for item in json_data['ilwidae_data']:
        for sangul in item['산출근거']:
            total_sangul += 1
            for field in empty_counts.keys():
                if not sangul.get(field, ''):
                    empty_counts[field] += 1
    
    print(f"총 산출근거 항목 수: {total_sangul}")
    for field, count in empty_counts.items():
        percentage = (count / total_sangul * 100) if total_sangul > 0 else 0
        if percentage > 50:
            print(f"  ⚠️  {field}: {count}개 비어있음 ({percentage:.1f}%)")
        else:
            print(f"  ✅ {field}: {count}개 비어있음 ({percentage:.1f}%)")
    
    # 데이터 타입 확인
    print("\n[수량 필드 데이터 타입 검사]")
    numeric_count = 0
    non_numeric_count = 0
    
    for item in json_data['ilwidae_data']:
        for sangul in item['산출근거']:
            if sangul.get('수량'):
                try:
                    float(sangul['수량'])
                    numeric_count += 1
                except:
                    non_numeric_count += 1
                    print(f"  ⚠️  숫자가 아닌 수량: {sangul['수량']} (행 {sangul['row_number']+1})")
    
    print(f"  숫자형 수량: {numeric_count}개")
    print(f"  비숫자형 수량: {non_numeric_count}개")

def main():
    """메인 실행 함수"""
    print("test1.xlsx 일위대가 추출 검증")
    print("=" * 70)
    
    # 1. 엑셀 데이터 읽기
    df, excel_hopyos = read_excel_data()
    
    # 2. JSON 데이터 읽기
    json_data, json_hopyos = read_json_data()
    
    # 3. 데이터 비교
    compare_data(excel_hopyos, json_hopyos, df, json_data)
    
    # 4. 데이터 완전성 검사
    check_data_completeness(df, json_data)
    
    # 최종 결과
    print("\n" + "=" * 70)
    print("검증 완료")
    print("=" * 70)
    
    # 요약
    print("\n[검증 요약]")
    print(f"• 엑셀 호표 수: {len(excel_hopyos)}개")
    print(f"• JSON 호표 수: {len(json_hopyos)}개")
    print(f"• 일치 여부: {'✅ 일치' if len(excel_hopyos) == len(json_hopyos) else '❌ 불일치'}")
    
    return excel_hopyos, json_hopyos

if __name__ == "__main__":
    main()