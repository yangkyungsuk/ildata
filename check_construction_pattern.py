import pandas as pd
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 건축구조내역.xlsx의 중기단가산출서 확인
df = pd.read_excel('건축구조내역.xlsx', sheet_name='중기단가산출서', header=None)

print("상위 20행 데이터:")
print("="*80)

for i in range(min(20, len(df))):
    row_data = []
    for j in range(min(10, len(df.columns))):
        cell = df.iloc[i, j]
        if pd.notna(cell):
            cell_str = str(cell).strip()
            if cell_str:
                row_data.append(f"{j}열: {cell_str}")
    if row_data:
        print(f"행 {i}: {' | '.join(row_data)}")