# Git 초기화 및 GitHub 업로드 가이드

## 1. GitHub 저장소 생성
1. https://github.com 에 로그인
2. 우측 상단 '+' 버튼 클릭 → 'New repository' 선택
3. Repository name: `korean-construction-excel-parser` (또는 원하는 이름)
4. Description: "한국 건설 견적서 Excel 파일 파서"
5. Public/Private 선택
6. **중요**: "Initialize this repository with:" 부분의 모든 체크박스는 **체크하지 마세요**
7. 'Create repository' 클릭

## 2. 로컬 Git 초기화
현재 프로젝트 폴더에서 다음 명령어를 실행하세요:

```bash
# Git 초기화
git init

# 모든 파일 추가
git add .

# 첫 커밋
git commit -m "Initial commit: Korean construction estimate Excel parser"
```

## 3. GitHub에 연결 및 업로드
GitHub에서 생성한 저장소 페이지에 나온 URL을 사용합니다:

```bash
# GitHub 저장소 연결 (YOUR_USERNAME을 실제 GitHub 사용자명으로 변경)
git remote add origin https://github.com/YOUR_USERNAME/korean-construction-excel-parser.git

# 기본 브랜치 이름 설정
git branch -M main

# GitHub에 푸시
git push -u origin main
```

## 4. 인증
- GitHub 사용자명과 비밀번호를 입력하라고 나올 수 있습니다
- 2021년 8월부터 GitHub는 비밀번호 대신 Personal Access Token 사용을 요구합니다
- Token 생성: GitHub → Settings → Developer settings → Personal access tokens → Generate new token

## 5. 업로드 확인
브라우저에서 GitHub 저장소 페이지를 새로고침하면 업로드된 파일들을 볼 수 있습니다.

## 문제 해결

### Permission denied 오류
```bash
git config --global user.name "Your Name"
git config --global user.email "your-email@example.com"
```

### 이미 Git 저장소가 있는 경우
```bash
# 기존 원격 저장소 확인
git remote -v

# 기존 원격 저장소 제거 (필요시)
git remote remove origin

# 새 원격 저장소 추가
git remote add origin https://github.com/YOUR_USERNAME/korean-construction-excel-parser.git
```