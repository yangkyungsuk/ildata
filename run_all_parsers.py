import subprocess
import sys
import io

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("="*60)
print("모든 파서 실행")
print("="*60)

# 파서 목록
parsers = [
    ('parser_test1.py', 'test1.xlsx'),
    ('parser_sgs.py', 'sgs.xls'),
    ('parser_construction.py', '건축구조내역2.xlsx'),
    ('parser_est.py', 'est.xlsx')
]

# 각 파서 실행
for parser_file, excel_file in parsers:
    print(f"\n{parser_file} 실행 중...")
    print("-" * 40)
    
    try:
        result = subprocess.run(
            ['python', parser_file],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            print(f"✓ {excel_file} 파싱 성공")
            # 마지막 몇 줄만 출력 (요약 정보)
            lines = result.stdout.strip().split('\n')
            for line in lines[-5:]:
                if '호표' in line or '항목' in line or '저장' in line:
                    print(f"  {line}")
        else:
            print(f"✗ {excel_file} 파싱 실패")
            print(f"  오류: {result.stderr}")
            
    except Exception as e:
        print(f"✗ {parser_file} 실행 오류: {str(e)}")

print("\n" + "="*60)
print("모든 파서 실행 완료")
print("="*60)