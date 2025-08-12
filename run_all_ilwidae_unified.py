"""
모든 일위대가 통합 파서 실행 스크립트
통합 JSON 구조로 모든 일위대가 데이터 파싱
"""
import os
import sys
import io
import json
from datetime import datetime

# UTF-8 인코딩 설정
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except:
    pass  # 이미 설정되어 있거나 가상환경에서 실행 중

def run_parser(parser_module, parser_class):
    """개별 파서 실행"""
    try:
        # 동적 import
        module = __import__(parser_module)
        parser_cls = getattr(module, parser_class)
        
        # 파서 인스턴스 생성 및 실행
        parser = parser_cls()
        output_file = parser.save_to_json()
        
        if output_file:
            print(f"✓ {parser.file_path} - {parser.sheet_name} 파싱 성공")
            return True, output_file
        else:
            print(f"✗ {parser.file_path} - {parser.sheet_name} 파싱 실패")
            return False, None
            
    except ImportError as e:
        print(f"✗ {parser_module} 모듈 로드 실패: {str(e)}")
        return False, None
    except Exception as e:
        print(f"✗ 파서 실행 중 오류: {str(e)}")
        return False, None

def check_file_exists(file_path):
    """파일 존재 여부 확인"""
    return os.path.exists(file_path)

def main():
    """메인 실행 함수"""
    print("=" * 70)
    print("일위대가 통합 파서 일괄 실행")
    print("=" * 70)
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("-" * 70)
    
    # 파서 목록 정의
    parsers = [
        {
            'name': 'ebs.xls - 일위대가_산근',
            'file': 'ebs.xls',
            'module': 'parser_ilwidae_ebs_unified',
            'class': 'IlwidaeEbsParser'
        },
        {
            'name': 'stmate.xlsx - 일위대가_산근',
            'file': 'stmate.xlsx',
            'module': 'parser_ilwidae_stmate_sanggun_unified',
            'class': 'IlwidaeStmateSanggunParser'
        },
        {
            'name': 'stmate.xlsx - 일위대가_호표',
            'file': 'stmate.xlsx',
            'module': 'parser_ilwidae_stmate_hopyo_unified',
            'class': 'IlwidaeStmateHopyoParser'
        }
    ]
    
    # 실행 결과 저장
    results = {
        'timestamp': datetime.now().isoformat(),
        'total': len(parsers),
        'success': 0,
        'failed': 0,
        'skipped': 0,
        'details': []
    }
    
    # 각 파서 실행
    for parser_info in parsers:
        print(f"\n[{parser_info['name']}]")
        
        # 파일 존재 확인
        if not check_file_exists(parser_info['file']):
            print(f"  ⚠ 파일이 존재하지 않음: {parser_info['file']}")
            results['skipped'] += 1
            results['details'].append({
                'parser': parser_info['name'],
                'status': 'skipped',
                'reason': 'file_not_found'
            })
            continue
        
        # 파서 실행
        success, output_file = run_parser(parser_info['module'], parser_info['class'])
        
        if success:
            results['success'] += 1
            results['details'].append({
                'parser': parser_info['name'],
                'status': 'success',
                'output': output_file
            })
        else:
            results['failed'] += 1
            results['details'].append({
                'parser': parser_info['name'],
                'status': 'failed'
            })
    
    # 결과 요약
    print("\n" + "=" * 70)
    print("실행 결과 요약")
    print("=" * 70)
    print(f"전체: {results['total']}개")
    print(f"성공: {results['success']}개")
    print(f"실패: {results['failed']}개")
    print(f"건너뜀: {results['skipped']}개")
    
    # 결과 파일 저장
    result_file = f"ilwidae_unified_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(result_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    print(f"\n실행 결과 저장: {result_file}")
    print("=" * 70)
    
    # 성공한 파일 목록 출력
    if results['success'] > 0:
        print("\n✅ 생성된 통합 JSON 파일:")
        for detail in results['details']:
            if detail['status'] == 'success':
                print(f"  - {detail['output']}")
    
    return results

if __name__ == "__main__":
    main()