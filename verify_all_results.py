import pandas as pd
import json
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def verify_test1():
    """test1.xlsx 검증"""
    print("\n" + "="*60)
    print("test1.xlsx 검증")
    print("="*60)
    
    # JSON 읽기
    with open('test1_parsed.json', 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    
    # 엑셀 읽기
    df = pd.read_excel('test1.xlsx', sheet_name='일위대가_산근', header=None)
    
    # 호표1 검증
    print("\n호표1 검증:")
    print(f"JSON 작업명: {json_data['hopyo_data']['호표1']['작업명']}")
    print(f"엑셀 3행 1열: {df.iloc[3, 1]}")  # 제1호표 행의 품명
    
    # 호표1 세부항목 검증
    print("\n호표1 세부항목:")
    for i, item in enumerate(json_data['hopyo_data']['호표1']['세부항목']):
        print(f"  {i+1}. {item['품명']} - {item['수량']} {item['단위']}")
    
    # 엑셀에서 확인
    print("\n엑셀 4-6행 (호표1 세부항목):")
    for i in range(4, 7):
        print(f"  행{i}: {df.iloc[i, 1]} - {df.iloc[i, 4]} {df.iloc[i, 3]}")
    
    print(f"\n총 호표 수: {json_data['hopyo_count']}")
    print(f"총 세부항목 수: {sum(data['항목수'] for data in json_data['hopyo_data'].values())}")

def verify_sgs():
    """sgs.xls 검증"""
    print("\n" + "="*60)
    print("sgs.xls 검증")
    print("="*60)
    
    # JSON 읽기
    with open('sgs_parsed.json', 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    
    # 엑셀 읽기
    df = pd.read_excel('sgs.xls', sheet_name='일위대가', header=None, engine='xlrd')
    
    # 호표1 검증
    print("\n호표1 검증:")
    print(f"JSON 작업명: {json_data['hopyo_data']['호표1']['작업명']}")
    print(f"엑셀 4행 0열: {df.iloc[4, 0]}")  # 제1호표 행
    
    # 호표1 세부항목 검증
    print("\n호표1 세부항목 (처음 3개):")
    for i, item in enumerate(json_data['hopyo_data']['호표1']['세부항목'][:3]):
        print(f"  {i+1}. {item['품명']} - {item['수량']} {item['단위']}")
    
    # 엑셀에서 확인
    print("\n엑셀 5-7행 (호표1 세부항목):")
    for i in range(5, 8):
        print(f"  행{i}: {df.iloc[i, 0]} - {df.iloc[i, 2]} {df.iloc[i, 3]}")
    
    # 호표1과 호표2 사이 데이터 확인
    print("\n호표1 세부항목 수:", len(json_data['hopyo_data']['호표1']['세부항목']))
    print("(계산 행 포함 여부 확인)")

def verify_construction():
    """건축구조내역2.xlsx 검증"""
    print("\n" + "="*60)
    print("건축구조내역2.xlsx 검증")
    print("="*60)
    
    # JSON 읽기
    with open('construction_parsed.json', 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    
    # 엑셀 읽기
    df = pd.read_excel('건축구조내역2.xlsx', sheet_name='일위대가', header=None)
    
    # 호표1 검증
    print("\n호표1 검증:")
    print(f"JSON 작업명: {json_data['hopyo_data']['호표1']['작업명']}")
    print(f"엑셀 3행 0열: {df.iloc[3, 0][:50]}...")  # 긴 텍스트 축약
    
    # 호표5 검증 (플랜지 접합)
    print("\n호표5 검증:")
    print(f"JSON 작업명: {json_data['hopyo_data']['호표5']['작업명']}")
    print(f"JSON 규격: {json_data['hopyo_data']['호표5']['규격']}")
    print(f"세부항목 수: {len(json_data['hopyo_data']['호표5']['세부항목'])}")
    
    # 세부항목 확인
    print("\n호표5 세부항목:")
    for item in json_data['hopyo_data']['호표5']['세부항목']:
        print(f"  - {item['품명']} ({item['수량']} {item['단위']})")

def verify_est():
    """est.xlsx 검증"""
    print("\n" + "="*60)
    print("est.xlsx 검증")
    print("="*60)
    
    # JSON 읽기
    with open('est_parsed.json', 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    
    # 엑셀 읽기
    df = pd.read_excel('est.xlsx', sheet_name='일위대가', header=None)
    
    # 호표1 검증
    print("\n호표1 검증:")
    print(f"JSON 작업명: {json_data['hopyo_data']['호표1']['작업명']}")
    print(f"엑셀 4행 2열: {df.iloc[4, 2]}")  # 제1호표 행의 품명
    
    # 세부항목 문제 확인
    print(f"\n호표1 세부항목 수: {len(json_data['hopyo_data']['호표1']['세부항목'])}")
    
    if len(json_data['hopyo_data']['호표1']['세부항목']) > 0:
        print("호표1 세부항목:")
        for item in json_data['hopyo_data']['호표1']['세부항목']:
            print(f"  - {item['품명']} ({item['수량']} {item['단위']})")
    
    # 엑셀에서 세부항목 확인
    print("\n엑셀 5-8행 (호표1 아래):")
    for i in range(5, 9):
        row_data = []
        for j in range(6):
            if pd.notna(df.iloc[i, j]):
                row_data.append(f"{j}:{df.iloc[i, j]}")
        print(f"  행{i}: {row_data[:4]}")
    
    print(f"\n전체 세부항목 수: {sum(data['항목수'] for data in json_data['hopyo_data'].values())}")
    print("(est는 세부항목이 적게 추출되는 문제 확인)")

def main():
    """모든 검증 실행"""
    verify_test1()
    verify_sgs()
    verify_construction()
    verify_est()
    
    print("\n" + "="*60)
    print("검증 완료")
    print("="*60)

if __name__ == "__main__":
    main()