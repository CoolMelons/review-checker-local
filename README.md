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

### 1️⃣ 다운로드
```bash
git clone https://github.com/YOUR-USERNAME/review-checker-local.git
cd review-checker-local
```

### 2️⃣ 설치
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
  김미정              2팀 /  8팀 ( 25.0%) - 평균 4.5점
    └ L                1팀 /  5팀 ( 20.0%)
    └ KK               1팀 /  2팀 ( 50.0%)
    └ TPC              2팀 /  6명 (검색 필요)

[Agency별 상세]
  L                5팀 / 68팀 (  7.4%) - 평균 5.0점
  KK               1팀 /  4팀 ( 25.0%) - 평균 5.0점
  GG               1팀 /  3팀 ( 33.3%) - 평균 5.0점
```
