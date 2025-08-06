import pandas as pd
import sys
import io
import re

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ebs.xls의 일위대가목록 확인
try:
    # 일위대가목록 읽기
    df_list = pd.read_excel('ebs.xls', sheet_name='일위대가목록', header=None, engine='xlrd')
    print("일위대가목록 시트 크기:", df_list.shape)
    print("\n[일위대가목록 내용]")
    
    # 호표 정보 추출
    ilwidae_items = []
    for i in range(len(df_list)):
        for j in range(len(df_list.columns)):
            cell = df_list.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                # 호표 패턴 찾기
                if '호표' in cell_str or re.match(r'^\d+$', cell_str):
                    # 같은 행의 정보 수집
                    row_info = []
                    for k in range(len(df_list.columns)):
                        c = df_list.iloc[i, k]
                        if pd.notna(c):
                            row_info.append(str(c).strip())
                    
                    if len(row_info) > 1:
                        ilwidae_items.append(row_info)
                        if len(ilwidae_items) <= 10:
                            print(f"행{i}: {' | '.join(row_info[:5])}")
    
    print(f"\n총 {len(ilwidae_items)}개 항목")
    
    # 일위대가_산근 시트에서 실제 일위대가 호표 찾기
    print("\n[일위대가_산근 시트에서 일위대가 호표 찾기]")
    
    df_sanggun = pd.read_excel('ebs.xls', sheet_name='일위대가_산근', header=None, engine='xlrd')
    
    # 일위대가 호표 패턴 (제N호표)
    ilwidae_pattern = r'제\s*(\d+)\s*호표'
    found_ilwidae = []
    
    for i in range(min(1000, len(df_sanggun))):  # 처음 1000행만 확인
        for j in range(min(5, len(df_sanggun.columns))):
            cell = df_sanggun.iloc[i, j]
            if pd.notna(cell):
                cell_str = str(cell).strip()
                match = re.match(ilwidae_pattern, cell_str)
                if match:
                    hopyo_num = match.group(1)
                    # 작업명 찾기
                    work_name = ''
                    for k in range(j+1, len(df_sanggun.columns)):
                        next_cell = df_sanggun.iloc[i, k]
                        if pd.notna(next_cell):
                            work_name = str(next_cell).strip()
                            if work_name:
                                break
                    
                    found_ilwidae.append({
                        'row': i,
                        'num': hopyo_num,
                        'work': work_name
                    })
                    
                    if len(found_ilwidae) <= 5:
                        print(f"  행{i}: 제{hopyo_num}호표 - {work_name}")
    
    if found_ilwidae:
        print(f"\n일위대가 호표 발견: {len(found_ilwidae)}개")
        print("→ ebs는 일위대가와 단가산출이 같은 시트에 혼재")
    else:
        print("\n일위대가 호표를 찾을 수 없음")
        print("→ 모두 단가산출 형식(#N)으로 되어 있음")
    
except Exception as e:
    print(f"오류: {str(e)}")
    import traceback
    traceback.print_exc()