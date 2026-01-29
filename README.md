# 📋 Review Checker - 데일리 성과제 리뷰 체크

**L (KLOOK), KK (KKDAY), GG (GetYourGuide)** 리뷰를 자동으로 조회하는 데스크톱 프로그램입니다.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac%20%7C%20Linux-lightgrey.svg)

---

## ✨ 주요 기능

- ✅ **L (KLOOK)** - 날짜별 리뷰 전체 수집 후 매칭
- ✅ **KK (KKDAY)** - 개별 예약번호 조회
- ✅ **GG (GetYourGuide)** - 날짜별 리뷰 전체 수집 후 매칭
- ✅ **가이드별 통계** - 가이드별 리뷰 현황 및 Agency 세부사항
- ✅ **Agency별 통계** - 에이전시별 리뷰 비율 및 평균 별점
- ✅ **개별 조회 필요 에이전시** - 기타 에이전시 목록 표시
- ✅ **실시간 진행률** - 조회 진행 상황 확인

---

## 🚀 빠른 시작

아래 두 가지 방법 중 편한 방식으로 실행하세요.  
- **방법 A (추천): Release에서 EXE 다운로드** → 설치/세팅 최소화
- **방법 B: ZIP 다운로드 후 Python으로 실행** → 개발환경에서 직접 실행

---

### ✅ 방법 A) Release에서 EXE 다운로드 (가장 쉬움 / 추천)

1️⃣ GitHub 저장소 상단 **Releases**로 이동  
2️⃣ 최신 버전에서 실행 파일 다운로드  
- Windows: `ReviewChecker.exe`

3️⃣ 다운로드한 파일 실행

> ⚠️ Windows에서 “알 수 없는 앱” 경고가 뜨면  
> **추가 정보 → 실행**으로 진행하세요.

---

### ✅ 방법 B) GitHub에서 ZIP 다운로드 후 Python으로 실행 (개발환경 필요)

#### 1️⃣ 다운로드
1) GitHub 페이지에서 **Code → Download ZIP** 클릭  
2) ZIP 압축 해제  
3) 압축 해제한 폴더로 이동

#### 2️⃣ 설치
```bash
pip install -r requirements.txt
```

### 3️⃣ 크롬 실행
```cmd
# Windows
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Chrome_debug_temp"

# Mac
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
```

### 4️⃣ 실행
```bash
python review_checker.py
```

---

## 📊 출력 예시

```
================================================================================
                              📈 전체 통계
================================================================================
👥 총 예약: 84팀 260명
   └ 리뷰 조회 대상: 75팀 238명 (L, KK, GG)
   └ 조회 제외: 9팀 22명 (TPC, D)

✓ 리뷰 확인: 7팀 / 75팀 (9.3%)
⭐ 평균 별점: 4.6점

[가이드별 상세]
  홍길동              2팀 /  8팀 ( 25.0%) - 평균 4.5점
    └ L                1팀 /  5팀 ( 20.0%)
    └ KK               1팀 /  2팀 ( 50.0%)
    └ TPC              2팀 /  6명 (검색 필요)

[Agency별 상세]
  L                5팀 / 68팀 (  7.4%) - 평균 5.0점
  KK               1팀 /  4팀 ( 25.0%) - 평균 5.0점
  GG               1팀 /  3팀 ( 33.3%) - 평균 5.0점
```
