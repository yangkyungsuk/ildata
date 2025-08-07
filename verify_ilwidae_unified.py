"""
일위대가 통합 JSON 구조 검증 시스템
품명, 규격, 단위, 수량 4개 컬럼 구조 검증
"""
import json
import os
import sys
import io
from typing import Dict, List, Tuple

# UTF-8 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

class IlwidaeUnifiedVerifier:
    """일위대가 통합 구조 검증 클래스"""
    
    def __init__(self):
        self.errors = []
        self.warnings = []
        
    def verify_json_structure(self, json_file: str) -> Tuple[bool, Dict]:
        """JSON 파일 구조 검증"""
        print(f"\n검증 중: {json_file}")
        print("-" * 50)
        
        # 파일 읽기
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except FileNotFoundError:
            self.errors.append(f"파일을 찾을 수 없음: {json_file}")
            return False, {}
        except json.JSONDecodeError as e:
            self.errors.append(f"JSON 파싱 오류: {str(e)}")
            return False, {}
        
        # 필수 최상위 필드 검증
        required_fields = ['file', 'sheet', 'total_ilwidae_count', 'validation', 'ilwidae_data']
        missing_fields = []
        
        for field in required_fields:
            if field not in data:
                missing_fields.append(field)
        
        if missing_fields:
            self.errors.append(f"누락된 필수 필드: {', '.join(missing_fields)}")
            return False, data
        
        # validation 구조 검증
        validation = data.get('validation', {})
        validation_fields = ['list_count', 'parsed_count', 'match']
        for field in validation_fields:
            if field not in validation:
                self.warnings.append(f"validation에 {field} 필드 누락")
        
        # ilwidae_data 검증
        ilwidae_data = data.get('ilwidae_data', [])
        if not isinstance(ilwidae_data, list):
            self.errors.append("ilwidae_data는 배열이어야 함")
            return False, data
        
        print(f"✓ 기본 구조 검증 완료")
        print(f"  - 파일: {data.get('file')}")
        print(f"  - 시트: {data.get('sheet')}")
        print(f"  - 일위대가 개수: {data.get('total_ilwidae_count')}")
        
        # 각 일위대가 항목 검증
        valid_items = 0
        invalid_items = 0
        
        for idx, item in enumerate(ilwidae_data):
            if self.verify_ilwidae_item(item, idx):
                valid_items += 1
            else:
                invalid_items += 1
        
        print(f"\n일위대가 항목 검증:")
        print(f"  - 유효: {valid_items}개")
        print(f"  - 무효: {invalid_items}개")
        
        # 검증 결과
        is_valid = len(self.errors) == 0
        
        return is_valid, data
    
    def verify_ilwidae_item(self, item: Dict, idx: int) -> bool:
        """개별 일위대가 항목 검증"""
        item_valid = True
        
        # 필수 필드 검증
        required = ['ilwidae_no', 'ilwidae_title', 'position', '산출근거']
        for field in required:
            if field not in item:
                self.errors.append(f"일위대가 {idx}: {field} 필드 누락")
                item_valid = False
        
        # ilwidae_title 구조 검증 (4개 컬럼)
        if 'ilwidae_title' in item:
            title = item['ilwidae_title']
            title_fields = ['품명', '규격', '단위', '수량']
            for field in title_fields:
                if field not in title:
                    self.warnings.append(f"일위대가 {item.get('ilwidae_no', idx)}: 타이틀에 {field} 누락")
        
        # position 구조 검증
        if 'position' in item:
            position = item['position']
            pos_fields = ['start_row', 'end_row', 'total_rows']
            for field in pos_fields:
                if field not in position:
                    self.warnings.append(f"일위대가 {item.get('ilwidae_no', idx)}: position에 {field} 누락")
        
        # 산출근거 항목 검증
        if '산출근거' in item:
            sangul_items = item['산출근거']
            if not isinstance(sangul_items, list):
                self.errors.append(f"일위대가 {item.get('ilwidae_no', idx)}: 산출근거는 배열이어야 함")
                item_valid = False
            else:
                # 각 산출근거 항목 검증
                for sangul_idx, sangul in enumerate(sangul_items):
                    if not self.verify_sangul_item(sangul, item.get('ilwidae_no', idx), sangul_idx):
                        item_valid = False
        
        return item_valid
    
    def verify_sangul_item(self, sangul: Dict, ilwidae_no: str, idx: int) -> bool:
        """산출근거 항목 검증 (4개 컬럼)"""
        sangul_valid = True
        
        # 필수 4개 컬럼
        required_fields = ['품명', '규격', '단위', '수량']
        missing = []
        
        for field in required_fields:
            if field not in sangul:
                missing.append(field)
        
        if missing:
            self.warnings.append(f"일위대가 {ilwidae_no} 산출근거 {idx}: {', '.join(missing)} 필드 누락")
        
        # row_number 확인
        if 'row_number' not in sangul:
            self.warnings.append(f"일위대가 {ilwidae_no} 산출근거 {idx}: row_number 누락")
        
        return sangul_valid
    
    def print_summary(self):
        """검증 결과 요약 출력"""
        print("\n" + "=" * 70)
        print("검증 결과 요약")
        print("=" * 70)
        
        if len(self.errors) == 0:
            print("✅ 오류 없음 - 구조가 올바름")
        else:
            print(f"❌ 오류 {len(self.errors)}개 발견:")
            for error in self.errors[:10]:  # 최대 10개만 표시
                print(f"  - {error}")
            if len(self.errors) > 10:
                print(f"  ... 외 {len(self.errors) - 10}개")
        
        if len(self.warnings) > 0:
            print(f"\n⚠️  경고 {len(self.warnings)}개:")
            for warning in self.warnings[:5]:  # 최대 5개만 표시
                print(f"  - {warning}")
            if len(self.warnings) > 5:
                print(f"  ... 외 {len(self.warnings) - 5}개")
        
        return len(self.errors) == 0
    
    def compare_with_original(self, unified_file: str, original_file: str):
        """기존 v4 구조와 비교"""
        print(f"\n기존 구조와 비교: {original_file}")
        print("-" * 50)
        
        try:
            with open(unified_file, 'r', encoding='utf-8') as f:
                unified = json.load(f)
            with open(original_file, 'r', encoding='utf-8') as f:
                original = json.load(f)
            
            # 기본 정보 비교
            print(f"파일명 일치: {unified.get('file') == original.get('file')}")
            print(f"시트명 일치: {unified.get('sheet') == original.get('sheet')}")
            
            # 호표 개수 비교
            unified_count = unified.get('total_ilwidae_count', 0)
            original_count = original.get('total_hopyo_count', 0)
            print(f"항목 개수: 통합({unified_count}) vs 기존({original_count})")
            
            # 데이터 손실 여부 확인
            if 'ilwidae_data' in unified and 'hopyo_data' in original:
                unified_items = len(unified['ilwidae_data'])
                original_items = len(original['hopyo_data'])
                
                if unified_items == original_items:
                    print(f"✓ 데이터 개수 일치 ({unified_items}개)")
                else:
                    print(f"✗ 데이터 개수 불일치: 통합({unified_items}) vs 기존({original_items})")
            
        except FileNotFoundError as e:
            print(f"파일을 찾을 수 없음: {e}")
        except Exception as e:
            print(f"비교 중 오류: {str(e)}")

def main():
    """메인 실행 함수"""
    print("=" * 70)
    print("일위대가 통합 JSON 구조 검증")
    print("=" * 70)
    
    verifier = IlwidaeUnifiedVerifier()
    
    # 검증할 파일 목록
    json_files = [
        'ebs_ilwidae_unified.json',
        'stmate_sanggun_ilwidae_unified.json', 
        'stmate_hopyo_ilwidae_unified.json'
    ]
    
    all_valid = True
    
    for json_file in json_files:
        if os.path.exists(json_file):
            verifier.errors = []  # 초기화
            verifier.warnings = []
            
            is_valid, data = verifier.verify_json_structure(json_file)
            
            if not is_valid:
                all_valid = False
        else:
            print(f"\n⚠ 파일이 존재하지 않음: {json_file}")
    
    # 전체 결과 요약
    verifier.print_summary()
    
    if all_valid:
        print("\n✅ 모든 파일이 통합 구조를 준수합니다.")
    else:
        print("\n❌ 일부 파일에서 구조 문제가 발견되었습니다.")
    
    return all_valid

if __name__ == "__main__":
    main()