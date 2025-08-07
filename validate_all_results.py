"""
모든 파싱 결과 검증 및 보고서 생성
"""
import json
import os
import pandas as pd

def validate_json_file(json_path):
    """JSON 파일 검증"""
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    stats = {
        'file': data['file'],
        'sheet': data['sheet'],
        'ilwidae_count': data['total_ilwidae_count'],
        'sangul_count': 0,
        'field_stats': {
            '품명': {'filled': 0, 'total': 0},
            '규격': {'filled': 0, 'total': 0},
            '단위': {'filled': 0, 'total': 0},
            '수량': {'filled': 0, 'total': 0}
        },
        'samples': []
    }
    
    # 산출근거 통계
    for item in data['ilwidae_data']:
        for sangul in item['산출근거']:
            stats['sangul_count'] += 1
            for field in ['품명', '규격', '단위', '수량']:
                stats['field_stats'][field]['total'] += 1
                if sangul.get(field, ''):
                    stats['field_stats'][field]['filled'] += 1
        
        # 샘플 수집
        if len(stats['samples']) < 3 and item['산출근거']:
            stats['samples'].append({
                'hopyo': item['ilwidae_no'],
                'title': item['ilwidae_title']['품명'],
                'first_item': item['산출근거'][0] if item['산출근거'] else None
            })
    
    return stats

def generate_report():
    """전체 검증 보고서 생성"""
    print("=" * 80)
    print("일위대가 파싱 결과 검증 보고서")
    print("=" * 80)
    
    json_files = [
        'result/test1_xlsx_ilwidae_unified.json',
        'result/est_xlsx_ilwidae_unified.json',
        'result/sgs_xls_ilwidae_unified.json',
        'result/건축구조내역_xlsx_ilwidae_unified.json',
        'result/건축구조내역2_xlsx_ilwidae_unified.json'
    ]
    
    all_stats = []
    
    for json_path in json_files:
        if os.path.exists(json_path):
            stats = validate_json_file(json_path)
            all_stats.append(stats)
            
            print(f"\n{'='*60}")
            print(f"파일: {stats['file']}")
            print(f"시트: {stats['sheet']}")
            print('='*60)
            
            print(f"\n[기본 정보]")
            print(f"  일위대가 개수: {stats['ilwidae_count']}개")
            print(f"  산출근거 항목: {stats['sangul_count']}개")
            
            print(f"\n[필드 채움율]")
            for field, field_stats in stats['field_stats'].items():
                if field_stats['total'] > 0:
                    rate = field_stats['filled'] / field_stats['total'] * 100
                    status = "✅" if rate >= 80 else "⚠️" if rate >= 50 else "❌"
                    print(f"  {status} {field}: {field_stats['filled']}/{field_stats['total']} ({rate:.1f}%)")
            
            print(f"\n[샘플 데이터]")
            for sample in stats['samples']:
                print(f"  제{sample['hopyo']}호표: {sample['title'][:30]}")
                if sample['first_item']:
                    item = sample['first_item']
                    print(f"    → {item['품명']} | {item['규격']} | {item['단위']} | {item['수량']}")
    
    # 전체 요약
    print("\n" + "=" * 80)
    print("전체 요약")
    print("=" * 80)
    
    total_ilwidae = sum(s['ilwidae_count'] for s in all_stats)
    total_sangul = sum(s['sangul_count'] for s in all_stats)
    
    print(f"\n총 파일 수: {len(all_stats)}개")
    print(f"총 일위대가: {total_ilwidae}개")
    print(f"총 산출근거: {total_sangul}개")
    
    # 파일별 성공률
    print(f"\n[파일별 파싱 성공률]")
    for stats in all_stats:
        success_rate = "✅ 성공" if stats['sangul_count'] > 0 else "❌ 실패"
        print(f"  {stats['file']}: {success_rate} (일위대가 {stats['ilwidae_count']}개, 산출근거 {stats['sangul_count']}개)")
    
    # 개선 필요 사항
    print(f"\n[개선 필요 사항]")
    for stats in all_stats:
        issues = []
        
        if stats['sangul_count'] == 0 and stats['ilwidae_count'] > 0:
            issues.append("산출근거 추출 실패")
        
        for field, field_stats in stats['field_stats'].items():
            if field_stats['total'] > 0:
                rate = field_stats['filled'] / field_stats['total'] * 100
                if rate < 50:
                    issues.append(f"{field} 필드 채움율 낮음 ({rate:.1f}%)")
        
        if issues:
            print(f"  {stats['file']}:")
            for issue in issues:
                print(f"    - {issue}")
    
    return all_stats

def check_original_excel():
    """원본 엑셀 파일과 비교"""
    print("\n" + "=" * 80)
    print("원본 엑셀 파일 검증")
    print("=" * 80)
    
    files = [
        ('test1.xlsx', '일위대가_산근'),
        ('est.xlsx', '일위대가'),
        ('sgs.xls', '일위대가'),
        ('건축구조내역.xlsx', '일위대가'),
        ('건축구조내역2.xlsx', '일위대가')
    ]
    
    for file_path, sheet_name in files:
        if os.path.exists(file_path):
            print(f"\n{file_path}:")
            try:
                if file_path.endswith('.xls'):
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
                else:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                print(f"  - 시트 크기: {df.shape[0]}행 x {df.shape[1]}열")
                
                # 데이터가 있는 행 수 계산
                non_empty_rows = 0
                for i in range(len(df)):
                    row_has_data = False
                    for j in range(min(10, len(df.columns))):
                        if pd.notna(df.iloc[i, j]):
                            row_has_data = True
                            break
                    if row_has_data:
                        non_empty_rows += 1
                
                print(f"  - 데이터 있는 행: {non_empty_rows}행")
                
            except Exception as e:
                print(f"  - 오류: {str(e)}")

if __name__ == "__main__":
    # 검증 보고서 생성
    stats = generate_report()
    
    # 원본 파일 확인
    check_original_excel()
    
    print("\n" + "=" * 80)
    print("검증 완료")
    print("=" * 80)