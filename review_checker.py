import os
import sys
import time
import pandas as pd
from datetime import datetime, timedelta
from tkinter import Tk, filedialog, Label, Button, Toplevel, StringVar, messagebox, Frame, Scrollbar, Canvas, Checkbutton, BooleanVar
from tkinter.ttk import Progressbar
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

REQUIRED_COLS = ["Date", "Area", "Product", "Agency", "Agency Code", "Main Guide", "People"]


class ReviewCheckerGUI:
    def __init__(self):
        self.root = Tk()
        self.root.title("📋 Review Checker")
        self.root.geometry("700x1200")  # 크기 증가
        
        self.driver = None
        self.df = None
        self.guide_groups = []  # [(date, product, guide), ...]
        self.guide_checkboxes = {}  # {(date, product, guide): BooleanVar}
        self.select_all_var = BooleanVar(value=True)
        
        self.klook_setup_done = False
        self.klook_current_date = None
        self.gg_setup_done = False
        self.gg_current_date = None
        
        # UI 구성
        self.setup_ui()
        
    def setup_ui(self):
        """UI 구성"""
        # 제목
        Label(self.root, text="📋 Review Checker", font=("Arial", 18, "bold")).pack(pady=15)
        
        # 1. 크롬 연결
        frame1 = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        frame1.pack(fill="x", padx=20, pady=5)
        
        Label(frame1, text="1️⃣ 크롬 연결 (디버그 모드)", font=("Arial", 12, "bold")).pack(anchor="w")
        Label(frame1, text="⚠️ L, KK, GG 로그인 필요", font=("Arial", 9), fg="red").pack(anchor="w")
        
        self.chrome_status = StringVar(value="🔴 크롬 미연결")
        Label(frame1, textvariable=self.chrome_status, font=("Arial", 10)).pack(anchor="w", pady=5)
        
        Button(frame1, text="🔌 크롬 연결", command=self.connect_chrome, 
               width=20, height=1, bg="#4CAF50", fg="white").pack(anchor="w")
        
        # 2. 엑셀 파일 선택
        frame2 = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        frame2.pack(fill="x", padx=20, pady=5)
        
        Label(frame2, text="2️⃣ 엑셀 파일 선택 (Excel for Guides)", font=("Arial", 12, "bold")).pack(anchor="w")
        
        self.file_status = StringVar(value="📁 파일 미선택")
        Label(frame2, textvariable=self.file_status, font=("Arial", 10)).pack(anchor="w", pady=5)
        
        Button(frame2, text="📁 파일 선택", command=self.select_file, 
               width=20, height=1, bg="#2196F3", fg="white").pack(anchor="w")
        
        # 3. 가이드 선택 (스크롤 가능)
        self.guide_frame = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        self.guide_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        Label(self.guide_frame, text="조회할 가이드 선택:", font=("Arial", 12, "bold")).pack(anchor="w")
        
        # 전체 선택 체크박스
        self.select_all_check = Checkbutton(
            self.guide_frame, 
            text="☑ 전체 선택", 
            variable=self.select_all_var,
            command=self.toggle_all
        )
        self.select_all_check.pack(anchor="w", pady=5)
        
        # 스크롤 가능한 가이드 리스트
        canvas_frame = Frame(self.guide_frame)
        canvas_frame.pack(fill="both", expand=True)
        
        self.canvas = Canvas(canvas_frame, height=200)
        scrollbar = Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 4. 시작 버튼
        Button(self.root, text="▶️ 선택한 가이드만 조회 시작", 
               command=self.start_processing, 
               width=30, height=2, 
               bg="#FF9800", fg="white",
               font=("Arial", 11, "bold")).pack(pady=10)
        
        # 5. 결과 표시 영역 (새로 추가!)
        result_frame = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        result_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        Label(result_frame, text="📊 조회 결과", font=("Arial", 12, "bold")).pack(anchor="w")
        
        # 스크롤 가능한 텍스트 영역
        result_scroll_frame = Frame(result_frame)
        result_scroll_frame.pack(fill="both", expand=True)
        
        from tkinter import Text
        result_scrollbar = Scrollbar(result_scroll_frame)
        result_scrollbar.pack(side="right", fill="y")
        
        self.result_text = Text(
            result_scroll_frame,
            height=10,
            width=60,
            yscrollcommand=result_scrollbar.set,
            font=("Consolas", 9),
            wrap="none"
        )
        self.result_text.pack(side="left", fill="both", expand=True)
        result_scrollbar.config(command=self.result_text.yview)
        
        # 진행 상황
        self.progress_var = StringVar(value="")
        Label(self.root, textvariable=self.progress_var, font=("Arial", 9)).pack(pady=5)
        
        # 버튼 프레임 (복사 + 종료)
        button_frame = Frame(self.root)
        button_frame.pack(pady=5)
        
        Button(button_frame, text="📋 Copy", 
               command=self.copy_results, width=20,
               bg="#9C27B0", fg="white").pack(side="left", padx=5)
        
        Button(button_frame, text="End", 
               command=self.quit_app, width=15).pack(side="left", padx=5)
        
    def connect_chrome(self):
        """크롬 연결"""
        try:
            options = Options()
            options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            self.driver = webdriver.Chrome(options=options)
            self.chrome_status.set("🟢 크롬 연결됨")
            messagebox.showinfo("성공", "크롬 연결 성공!\n\nL, KK, GG에 로그인했는지 확인하세요.")
        except Exception as e:
            self.chrome_status.set("🔴 크롬 연결 실패")
            messagebox.showerror("연결 실패", 
                f"크롬 연결 실패: {e}\n\n다음 명령어로 크롬을 실행하세요:\n\n"
                'Windows:\n"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" '
                '--remote-debugging-port=9222 --user-data-dir="C:\\Chrome_debug_temp"\n\n'
                'Mac:\n/Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome '
                '--remote-debugging-port=9222')
    
    def select_file(self):
        """엑셀 파일 선택"""
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!")
            return
        
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # 엑셀 읽기
            df = pd.read_excel(file_path)
            df = self.normalize_columns(df)
            
            # 서울 필터링
            df = df[df["Area"].str.strip().str.lower() == "seoul"].copy()
            
            # 데이터 준비
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Agency"] = df["Agency"].astype(str).str.strip()
            df["Agency Code"] = df["Agency Code"].astype(str).str.strip()
            
            self.df = df
            self.file_status.set(f"✅ 파일 로드 완료: {len(df)}개 예약")
            
            # 가이드 그룹 추출 및 표시
            self.extract_and_display_guides()
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")
    
    def extract_and_display_guides(self):
        """가이드 그룹 추출 및 체크박스 표시"""
        # 기존 체크박스 제거
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.guide_groups = []
        self.guide_checkboxes = {}
        
        # 날짜, 투어, 가이드로 그룹화
        grouped = self.df.groupby(['Date', 'Product', 'Main Guide'])
        
        for (date_val, product, guide), group in grouped:
            self.guide_groups.append((date_val, product, guide))
            
            # 체크박스 변수
            var = BooleanVar(value=True)  # 기본 전체 선택
            self.guide_checkboxes[(date_val, product, guide)] = var
            
            # 체크박스 생성
            team_count = len(group)
            people_count = group['People'].sum() if 'People' in group.columns else 0
            
            label_text = f"{date_val.strftime('%Y-%m-%d')} | {product} | {guide} ({team_count}팀, {people_count}명)"
            
            from tkinter.ttk import Checkbutton as TtkCheckbutton
            cb = TtkCheckbutton(
                self.scrollable_frame,
                text=label_text,
                variable=var
            )
            cb.pack(anchor="w", padx=5, pady=2)
        
        messagebox.showinfo("완료", f"{len(self.guide_groups)}개 가이드 그룹을 찾았습니다.")
    
    def display_results(self, stats):
        """결과를 UI에 표시"""
        self.result_text.delete(1.0, "end")  # 기존 내용 삭제
        
        result = []
        result.append("=" * 60)
        result.append("📈 전체 통계")
        result.append("=" * 60)
        result.append(f"👥 총 예약: {stats['total_teams']}팀 {stats['total_people']}명")
        
        # 조회 대상
        reviewed_agencies = [a for a in ['L', 'KK', 'GG'] if stats['agencies'][a]['total'] > 0]
        result.append(f"   └ 리뷰 조회 대상: {stats['reviewed_total']}팀 {stats['reviewed_people']}명 ({', '.join(reviewed_agencies)})")
        
        other_total = stats['total_teams'] - stats['reviewed_total']
        other_people = stats['total_people'] - stats['reviewed_people']
        if other_total > 0:
            other_agencies = list(stats['other_agencies'].keys())
            result.append(f"   └ 조회 제외: {other_total}팀 {other_people}명 ({', '.join(other_agencies)})")
        
        if stats['reviewed_total'] > 0:
            pct = (stats['total_checked'] / stats['reviewed_total']) * 100
            result.append(f"\n✓ 리뷰 확인: {stats['total_checked']}팀 / {stats['reviewed_total']}팀 ({pct:.1f}%)")
        
        if stats['total_ratings']:
            avg_all = sum(stats['total_ratings']) / len(stats['total_ratings'])
            result.append(f"⭐ 평균 별점: {avg_all:.1f}점\n")
        else:
            result.append("⭐ 평균 별점: N/A\n")
        
        # 가이드별 상세
        result.append("\n[가이드별 상세]")
        result.append("-" * 60)
        for guide_name, guide_stat in stats['guides'].items():
            if guide_stat['total'] > 0:
                pct = (guide_stat['checked'] / guide_stat['total']) * 100
                avg = sum(guide_stat['ratings']) / len(guide_stat['ratings']) if guide_stat['ratings'] else 0
                line = f"  {guide_name:15} {guide_stat['checked']:2}팀 / {guide_stat['total']:2}팀 ({pct:5.1f}%)"
                if avg > 0:
                    line += f" - 평균 {avg:.1f}점"
                result.append(line)
                
                # Agency 세부
                for agency_code in ['L', 'KK', 'GG']:
                    agency_stat = guide_stat['agencies'][agency_code]
                    if agency_stat['total'] > 0:
                        agency_pct = (agency_stat['checked'] / agency_stat['total']) * 100
                        agency_avg = sum(agency_stat['ratings']) / len(agency_stat['ratings']) if agency_stat['ratings'] else 0
                        line = f"    └ {agency_code:15} {agency_stat['checked']:2}팀 / {agency_stat['total']:2}팀 ({agency_pct:5.1f}%)"
                        if agency_avg > 0:
                            line += f" - 평균 {agency_avg:.1f}점"
                        result.append(line)
                
                # 기타 에이전시
                for other_agency, bookings in guide_stat['other_agencies'].items():
                    if len(bookings) > 0:
                        total_people = sum(b['people'] for b in bookings)
                        result.append(f"    └ {other_agency:15} {len(bookings):2}팀 / {total_people:3}명 (검색 필요)")
        
        # Agency별 상세
        result.append("\n[Agency별 상세]")
        result.append("-" * 60)
        for agency_code, agency_stat in stats['agencies'].items():
            if agency_stat['total'] > 0:
                pct = (agency_stat['checked'] / agency_stat['total']) * 100
                avg = sum(agency_stat['ratings']) / len(agency_stat['ratings']) if agency_stat['ratings'] else 0
                line = f"  {agency_code:15} {agency_stat['checked']:2}팀 / {agency_stat['total']:2}팀 ({pct:5.1f}%)"
                if avg > 0:
                    line += f" - 평균 {avg:.1f}점"
                result.append(line)
        
        # 개별 조회 필요 에이전시
        if stats['other_agencies']:
            result.append("\n[개별 조회 필요 에이전시]")
            result.append("-" * 60)
            for agency_code, agency_data in stats['other_agencies'].items():
                result.append(f"  {agency_code:15} {agency_data['total']:2}팀")
                for booking in agency_data['bookings']:
                    result.append(f"    · {booking['code']} ({booking['guide']})")
        
        result.append("\n" + "=" * 60)
        
        # UI에 표시
        self.result_text.insert("end", "\n".join(result))
    
    def toggle_all(self):
        """전체 선택/해제"""
        select_all = self.select_all_var.get()
        for var in self.guide_checkboxes.values():
            var.set(select_all)
    
    def start_processing(self):
        """선택한 가이드만 조회"""
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!")
            return
        
        if self.df is None:
            messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
            return
        
        # 선택된 가이드 확인
        selected_guides = [
            key for key, var in self.guide_checkboxes.items() if var.get()
        ]
        
        if not selected_guides:
            messagebox.showerror("오류", "최소 1개 이상의 가이드를 선택하세요!")
            return
        
        # 선택된 가이드의 데이터만 필터링
        filtered_df = pd.DataFrame()
        for date_val, product, guide in selected_guides:
            mask = (
                (self.df['Date'] == date_val) & 
                (self.df['Product'] == product) & 
                (self.df['Main Guide'] == guide)
            )
            filtered_df = pd.concat([filtered_df, self.df[mask]])
        
        # 기존 select_file_and_start 로직 실행 (filtered_df 사용)
        self.select_file_and_start(filtered_df)
    
    def select_file_and_start(self, df=None):
        """엑셀 파일 선택 후 처리 시작 (또는 필터링된 df 처리)"""
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!")
            return
        
        # df가 없으면 파일 선택 (레거시)
        if df is None:
            file_path = filedialog.askopenfilename(
                title="엑셀 파일 선택",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not file_path:
                return
            
            self.progress_var.set("파일 처리 중...")
            self.root.update()
            
            try:
                # 엑셀 읽기
                df = pd.read_excel(file_path)
                df = self.normalize_columns(df)
                
                # 서울 필터링
                df = df[df["Area"].str.strip().str.lower() == "seoul"].copy()
                
                # 데이터 준비
                df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
                df["Agency"] = df["Agency"].astype(str).str.strip()
                df["Agency Code"] = df["Agency Code"].astype(str).str.strip()
                
            except Exception as e:
                messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")
                return
        
        # 여기서부터 기존 로직 (df 사용)
        try:
            
            # 결과 컬럼 추가
            df["Review_Status"] = ""
            df["Rating"] = ""
            df["Check"] = ""
            
            # KLOOK, GG 설정 초기화
            self.klook_setup_done = False
            self.klook_current_date = None
            self.gg_setup_done = False
            self.gg_current_date = None
            
            # 통계 초기화
            stats = {
                'total_teams': 0,
                'total_people': 0,
                'total_checked': 0,
                'total_ratings': [],
                'agencies': {
                    'L': {'name': 'KLOOK', 'total': 0, 'checked': 0, 'ratings': []},
                    'KK': {'name': 'KKDAY', 'total': 0, 'checked': 0, 'ratings': []},
                    'GG': {'name': 'GetYourGuide', 'total': 0, 'checked': 0, 'ratings': []}
                },
                'guides': {},  # {guide_name: {total, checked, ratings, agencies: {L: {}, KK: {}, GG: {}}}}
                'other_agencies': {},  # {agency_code: {name, total, people, bookings: [{code, guide, people}]}}
                'reviewed_total': 0,  # L+KK+GG 합계
                'reviewed_people': 0
            }
            
            # 진행창 생성
            progress_window = self.create_progress_window()
            progress_bar = progress_window.progress_bar
            progress_label = progress_window.label
            
            print("\n" + "="*80)
            print("📊 리뷰 조회 시작".center(80))
            print("="*80 + "\n")
            
            # 날짜별로 리뷰 수집
            unique_dates = df['Date'].unique()
            all_reviews = {
                'L': {},   # {date: {code: rating}}
                'KK': {},
                'GG': {}
            }
            
            print("=" * 80)
            print("1단계: 날짜별 리뷰 수집")
            print("=" * 80)
            
            for date_val in unique_dates:
                print(f"\n📅 {pd.to_datetime(date_val).strftime('%Y-%m-%d')}")
                print("-" * 60)
                
                # KLOOK 수집
                klook_reviews = self.collect_klook_reviews(date_val)
                all_reviews['L'][date_val] = klook_reviews
                
                # KKDAY는 개별 조회
                all_reviews['KK'][date_val] = {}
                
                # GG 수집
                gg_reviews = self.collect_gg_reviews(date_val)
                all_reviews['GG'][date_val] = gg_reviews
            
            print("\n" + "=" * 80)
            print("2단계: 예약번호 매칭 및 출력")
            print("=" * 80)
            
            # 날짜 → 투어 → 가이드별로 그룹화하여 처리
            grouped = df.groupby(['Date', 'Product', 'Main Guide'])
            processed_count = 0
            total = len(df)
            
            current_date = None
            
            for (date_val, product, guide), group in grouped:
                # 날짜가 바뀌면 날짜 헤더 출력
                if current_date != date_val:
                    if current_date is not None:
                        print()
                    print(f"\n{'='*80}")
                    print(f"📅 {date_val.strftime('%Y-%m-%d (%A)')}")
                    print(f"{'='*80}\n")
                    current_date = date_val
                
                # 투어/가이드별 정보
                people_count = group['People'].sum() if 'People' in group.columns else 0
                team_count = len(group)
                
                print(f"📍 투어: {product}")
                print(f"👤 가이드: {guide}")
                print(f"👥 총: {team_count}팀 {people_count}명\n")
                
                stats['total_teams'] += team_count
                stats['total_people'] += people_count
                
                # 가이드별 통계 초기화
                if guide not in stats['guides']:
                    stats['guides'][guide] = {
                        'total': 0, 
                        'checked': 0, 
                        'ratings': [],
                        'agencies': {
                            'L': {'total': 0, 'checked': 0, 'ratings': []},
                            'KK': {'total': 0, 'checked': 0, 'ratings': []},
                            'GG': {'total': 0, 'checked': 0, 'ratings': []}
                        },
                        'other_agencies': {}  # {agency_code: [{code, people}]}
                    }
                stats['guides'][guide]['total'] += team_count
                
                # Agency별 처리
                for agency in ['L', 'KK', 'GG']:
                    agency_group = group[group['Agency'] == agency]
                    if len(agency_group) == 0:
                        continue
                    
                    print(f"[{agency}]")
                    print("-" * 60)
                    
                    # 현재 그룹의 체크 카운트
                    current_checked = 0
                    current_ratings = []
                    
                    for idx, row in agency_group.iterrows():
                        code = row["Agency Code"]
                        date = row["Date"]
                        people = row.get("People", 0)
                        
                        processed_count += 1
                        progress_label.config(text=f"매칭 중: {processed_count}/{total} - {agency} {code}")
                        progress_bar["value"] = (processed_count / total) * 100
                        progress_window.window.update()
                        
                        # 수집된 데이터에서 매칭
                        status = "NO"
                        rating = ""
                        
                        if agency == "L" or agency == "GG":
                            # KLOOK, GG는 수집된 데이터에서 찾기
                            date_reviews = all_reviews[agency].get(date, {})
                            if code in date_reviews:
                                status = "YES"
                                rating = date_reviews[code]
                        elif agency == "KK":
                            # KKDAY는 개별 조회 (기존 방식 유지)
                            status, rating = self.check_kkday(code, date)
                        else:
                            status = "SKIP"
                        
                        # 결과 저장
                        df.at[idx, "Review_Status"] = status
                        df.at[idx, "Rating"] = rating
                        
                        # 가이드-Agency 통계 카운트
                        stats['guides'][guide]['agencies'][agency]['total'] += 1
                        stats['agencies'][agency]['total'] += 1
                        stats['reviewed_total'] += 1
                        stats['reviewed_people'] += people
                        
                        # 체크 표시 및 통계
                        if status == "YES":
                            df.at[idx, "Check"] = "✓"
                            stats['agencies'][agency]['checked'] += 1
                            stats['total_checked'] += 1
                            stats['guides'][guide]['checked'] += 1
                            stats['guides'][guide]['agencies'][agency]['checked'] += 1
                            current_checked += 1
                            
                            if rating and rating.replace('.', '').isdigit():
                                rating_val = float(rating)
                                stats['agencies'][agency]['ratings'].append(rating_val)
                                stats['total_ratings'].append(rating_val)
                                stats['guides'][guide]['ratings'].append(rating_val)
                                stats['guides'][guide]['agencies'][agency]['ratings'].append(rating_val)
                                current_ratings.append(rating_val)
                                print(f"  ✓ {code} ({rating}점)")
                            else:
                                print(f"  ✓ {code}")
                        else:
                            df.at[idx, "Check"] = "✗"
                            print(f"  ✗ {code}")
                        
                        time.sleep(0.3)
                    
                    # Agency별 요약 (현재 그룹만)
                    current_total = len(agency_group)
                    if current_total > 0:
                        pct = (current_checked / current_total) * 100
                        avg = sum(current_ratings) / len(current_ratings) if current_ratings else 0
                        print(f"\n  📊 {current_checked}/{current_total}팀 ({pct:.1f}%)", end="")
                        if avg > 0:
                            print(f" - 평균 {avg:.1f}점\n")
                        else:
                            print("\n")
                
                # 기타 Agency 처리 (조회 안 함)
                other_group = group[~group['Agency'].isin(['L', 'KK', 'GG'])]
                for idx, row in other_group.iterrows():
                    agency = row["Agency"]
                    code = row["Agency Code"]
                    people = row.get("People", 0)
                    
                    # 전체 통계에 기타 에이전시 추가
                    if agency not in stats['other_agencies']:
                        stats['other_agencies'][agency] = {
                            'name': agency,
                            'total': 0,
                            'people': 0,
                            'bookings': []
                        }
                    
                    stats['other_agencies'][agency]['total'] += 1
                    stats['other_agencies'][agency]['people'] += people
                    stats['other_agencies'][agency]['bookings'].append({
                        'code': code,
                        'guide': guide,
                        'people': people
                    })
                    
                    # 가이드별 기타 에이전시 추가
                    if agency not in stats['guides'][guide]['other_agencies']:
                        stats['guides'][guide]['other_agencies'][agency] = []
                    
                    stats['guides'][guide]['other_agencies'][agency].append({
                        'code': code,
                        'people': people
                    })
            
            # 전체 통계
            print(f"\n{'='*80}")
            print("📈 전체 통계".center(80))
            print(f"{'='*80}\n")
            print(f"👥 총 예약: {stats['total_teams']}팀 {stats['total_people']}명")
            
            # 조회 대상 에이전시 표시
            reviewed_agencies = []
            for agency_code in ['L', 'KK', 'GG']:
                if stats['agencies'][agency_code]['total'] > 0:
                    reviewed_agencies.append(agency_code)
            
            print(f"   └ 리뷰 조회 대상: {stats['reviewed_total']}팀 {stats['reviewed_people']}명 ({', '.join(reviewed_agencies)})")
            
            # 조회 제외 에이전시 표시
            other_total = stats['total_teams'] - stats['reviewed_total']
            other_people = stats['total_people'] - stats['reviewed_people']
            if other_total > 0:
                other_agencies = list(stats['other_agencies'].keys())
                print(f"   └ 조회 제외: {other_total}팀 {other_people}명 ({', '.join(other_agencies)})")
            
            if stats['reviewed_total'] > 0:
                pct = (stats['total_checked'] / stats['reviewed_total']) * 100
                print(f"\n✓ 리뷰 확인: {stats['total_checked']}팀 / {stats['reviewed_total']}팀 ({pct:.1f}%)")
            
            if stats['total_ratings']:
                avg_all = sum(stats['total_ratings']) / len(stats['total_ratings'])
                print(f"⭐ 평균 별점: {avg_all:.1f}점\n")
            else:
                print(f"⭐ 평균 별점: N/A\n")
            
            print("[가이드별 상세]")
            print("-" * 60)
            agency_names = {'L': 'L', 'KK': 'KK', 'GG': 'GG'}
            for guide_name, guide_stat in stats['guides'].items():
                if guide_stat['total'] > 0:
                    pct = (guide_stat['checked'] / guide_stat['total']) * 100
                    avg = sum(guide_stat['ratings']) / len(guide_stat['ratings']) if guide_stat['ratings'] else 0
                    print(f"  {guide_name:15} {guide_stat['checked']:2}팀 / {guide_stat['total']:2}팀 ({pct:5.1f}%)", end="")
                    if avg > 0:
                        print(f" - 평균 {avg:.1f}점")
                    else:
                        print()
                    
                    # 가이드별 Agency 세부내역
                    for agency_code, agency_name in agency_names.items():
                        agency_stat = guide_stat['agencies'][agency_code]
                        if agency_stat['total'] > 0:
                            agency_pct = (agency_stat['checked'] / agency_stat['total']) * 100
                            agency_avg = sum(agency_stat['ratings']) / len(agency_stat['ratings']) if agency_stat['ratings'] else 0
                            print(f"    └ {agency_name:15} {agency_stat['checked']:2}팀 / {agency_stat['total']:2}팀 ({agency_pct:5.1f}%)", end="")
                            if agency_avg > 0:
                                print(f" - 평균 {agency_avg:.1f}점")
                            else:
                                print()
                    
                    # 기타 에이전시
                    for other_agency, bookings in guide_stat['other_agencies'].items():
                        if len(bookings) > 0:
                            total_people = sum(b['people'] for b in bookings)
                            print(f"    └ {other_agency:15} {len(bookings):2}팀 / {total_people:3}명 (검색 필요)")
            
            print()
            print("[Agency별 상세]")
            print("-" * 60)
            for agency_code, agency_stat in stats['agencies'].items():
                if agency_stat['total'] > 0:
                    pct = (agency_stat['checked'] / agency_stat['total']) * 100
                    avg = sum(agency_stat['ratings']) / len(agency_stat['ratings']) if agency_stat['ratings'] else 0
                    print(f"  {agency_code:15} {agency_stat['checked']:2}팀 / {agency_stat['total']:2}팀 ({pct:5.1f}%)", end="")
                    if avg > 0:
                        print(f" - 평균 {avg:.1f}점")
                    else:
                        print()
            
            # 개별 조회 필요 에이전시
            if stats['other_agencies']:
                print()
                print("[개별 조회 필요 에이전시]")
                print("-" * 60)
                for agency_code, agency_data in stats['other_agencies'].items():
                    print(f"  {agency_code:15} {agency_data['total']:2}팀")
                    for booking in agency_data['bookings']:
                        print(f"    · {booking['code']} ({booking['guide']})")
            
            print(f"\n{'='*80}\n")
            
            progress_window.window.destroy()
            
            # UI에 결과 표시
            self.display_results(stats)
            
            # 최종 메시지
            if stats['total_ratings']:
                final_msg = f"✅ 완료!\n\n리뷰 확인: {stats['total_checked']}/{stats['reviewed_total']}팀 ({stats['total_checked']/stats['reviewed_total']*100:.1f}%)\n평균 별점: {sum(stats['total_ratings'])/len(stats['total_ratings']):.1f}점"
            else:
                final_msg = f"✅ 완료!\n\n리뷰 확인: {stats['total_checked']}/{stats['reviewed_total']}팀"
            
            self.progress_var.set("✅ 완료!")
            messagebox.showinfo("완료", final_msg)
            print(f"✅ 조회 완료!\n")
            
        except Exception as e:
            self.progress_var.set(f"❌ 오류: {str(e)}")
            print(f"오류 발생: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"처리 중 오류 발생:\n{e}")
    
    def create_progress_window(self):
        """진행바 창 생성"""
        window = Toplevel(self.root)
        window.title("처리 중...")
        window.geometry("400x100")
        
        label = Label(window, text="시작 중...", font=("Arial", 10))
        label.pack(pady=10)
        
        progress_bar = Progressbar(window, length=350, mode="determinate")
        progress_bar.pack(pady=10)
        
        # 창 닫기 방지
        window.protocol("WM_DELETE_WINDOW", lambda: None)
        
        # 객체에 참조 저장
        window.progress_bar = progress_bar
        window.label = label
        window.window = window
        
        return window
    
    def normalize_columns(self, df):
        """컬럼 정규화"""
        df.columns = [str(c).strip() for c in df.columns]
        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(f"필수 컬럼 누락: {missing}")
        return df
    
    
    def collect_klook_reviews(self, date):
        """KLOOK에서 해당 날짜의 모든 리뷰 수집"""
        reviews = {}  # {booking_code: rating}
        
        try:
            print(f"\n🔍 KLOOK 리뷰 수집 중... (날짜: {date.strftime('%Y-%m-%d')})")
            
            # KLOOK 페이지로 이동
            self.driver.get("https://merchant.klook.com/reviews")
            time.sleep(2)
            
            # 날짜 필터 설정
            try:
                date_str = date.strftime("%Y-%m-%d")
                
                # Product 필터 선택
                product_dropdown = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[1]/form[2]/div[1]/div[2]/div/span'
                )
                product_dropdown.click()
                time.sleep(1)
                
                participation_options = self.driver.find_elements(
                    By.XPATH,
                    '//li[contains(text(), "Participation time")]'
                )
                for opt in participation_options:
                    if "Participation time" in opt.text:
                        opt.click()
                        time.sleep(0.5)
                        break
                
                # 날짜 입력
                from selenium.webdriver.common.keys import Keys
                
                main_input = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[1]/form[2]/div[2]/div[2]/div/span/span/span/input[1]'
                )
                main_input.click()
                time.sleep(1)
                
                popup_start_input = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div/div/div[1]/div[1]/div[1]/div/input'))
                )
                popup_start_input.click()
                popup_start_input.send_keys(Keys.CONTROL + 'a')
                popup_start_input.send_keys(date_str)
                time.sleep(0.3)
                
                popup_end_input = self.driver.find_element(
                    By.XPATH,
                    '/html/body/div[3]/div/div/div/div/div[1]/div[2]/div[1]/div/input'
                )
                popup_end_input.click()
                popup_end_input.send_keys(Keys.CONTROL + 'a')
                popup_end_input.send_keys(date_str)
                time.sleep(0.3)
                
                # Search 버튼
                search_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[2]/button[1]'))
                )
                search_btn.click()
                time.sleep(3)
                
            except Exception as e:
                print(f"  ⚠ 날짜 필터 설정 실패: {e}")
            
            # 모든 페이지 순회하며 수집
            page_num = 1
            while page_num <= 20:
                try:
                    # 전체 리뷰 리스트에서 Booking reference ID와 Stars 추출
                    # 방법 1: 테이블 행으로 읽기
                    rows = self.driver.find_elements(
                        By.XPATH,
                        '//*[@id="klook-content"]/div/div[2]/div/div/div/div/div/div/div/div/div/div/table/tbody/tr'
                    )
                    
                    for row in rows:
                        try:
                            # Booking reference ID (첫 번째 열)
                            code = row.find_element(By.XPATH, './td[1]/a').text.strip()
                            # Stars (6번째 열)
                            rating_text = row.find_element(By.XPATH, './td[6]').text.strip()
                            
                            if code:
                                reviews[code] = rating_text if rating_text.isdigit() else ""
                        except:
                            continue
                    
                    print(f"  → 페이지 {page_num}: {len(rows)}개 리뷰")
                    
                    # 다음 페이지 버튼
                    try:
                        next_btn = self.driver.find_element(
                            By.XPATH,
                            '//li[contains(@class, "ant-pagination-next") and not(contains(@class, "ant-pagination-disabled"))]/a'
                        )
                        next_btn.click()
                        time.sleep(2)
                        page_num += 1
                    except:
                        break
                        
                except Exception as e:
                    break
            
            print(f"  ✓ KLOOK: {len(reviews)}개 리뷰 수집 완료")
            return reviews
            
        except Exception as e:
            print(f"  ✗ KLOOK 수집 실패: {e}")
            import traceback
            traceback.print_exc()
            return reviews
    
    def collect_kkday_reviews(self, date):
        """KKDAY에서 해당 날짜의 모든 리뷰 수집"""
        reviews = {}
        
        try:
            print(f"\n🔍 KKDAY 리뷰 수집 중... (날짜: {date.strftime('%Y-%m-%d')})")
            
            # KKDAY는 개별 조회만 가능하므로 빈 딕셔너리 반환
            # 실제로는 예약 코드를 하나씩 조회해야 함
            print(f"  ⚠ KKDAY는 개별 조회 방식 유지")
            return reviews
            
        except Exception as e:
            print(f"  ✗ KKDAY 수집 실패: {e}")
            return reviews
    
    def collect_gg_reviews(self, date):
        """GG에서 해당 날짜의 모든 리뷰 수집"""
        reviews = {}
        
        try:
            from datetime import timedelta
            
            print(f"\n🔍 GG 리뷰 수집 중... (날짜: {date.strftime('%Y-%m-%d')})")
            
            # GG 페이지로 이동
            self.driver.get("https://supplier.getyourguide.com/performance/reviews")
            time.sleep(3)
            
            # More Filters 클릭
            try:
                more_filters = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="__nuxt"]/div/div/main/div[1]/div/div[2]/div[1]/div/div[3]/button'))
                )
                more_filters.click()
                time.sleep(1)
            except:
                pass
            
            # 날짜 선택
            try:
                prev_day = date - timedelta(days=1)
                prev_day_num = prev_day.day
                curr_day_num = date.day
                
                calendar_btn = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="date-range"]/span/span/span'
                )
                calendar_btn.click()
                time.sleep(1)
                
                prev_day_cell = self.driver.find_element(
                    By.XPATH,
                    f'//span[@class="p-datepicker-day" and text()="{prev_day_num}"]'
                )
                prev_day_cell.click()
                time.sleep(0.3)
                
                curr_day_cell = self.driver.find_element(
                    By.XPATH,
                    f'//span[@class="p-datepicker-day" and text()="{curr_day_num}"]'
                )
                curr_day_cell.click()
                time.sleep(5)  # 결과 로딩 대기
                
            except Exception as e:
                print(f"  ⚠ 날짜 선택 실패: {e}")
            
            # 모든 페이지 순회
            page_num = 1
            while page_num <= 10:
                try:
                    # Show details 모두 열기
                    show_buttons = self.driver.find_elements(By.XPATH, '//button[contains(., "Show details")]')
                    for btn in show_buttons:
                        try:
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                            time.sleep(0.2)
                            btn.click()
                            time.sleep(0.3)
                        except:
                            continue
                    
                    # 모든 예약번호와 별점 수집
                    booking_elems = self.driver.find_elements(
                        By.XPATH,
                        '//a[contains(@href, "booking") or contains(text(), "GYG")]'
                    )
                    
                    print(f"  → 페이지 {page_num}: {len(booking_elems)}개 예약번호 발견")
                    
                    for elem in booking_elems:
                        try:
                            code = elem.text.strip()
                            if code.startswith("GYG"):
                                # 별점 찾기
                                try:
                                    parent = elem.find_element(By.XPATH, './ancestor::div[contains(@class, "c-review") or contains(@class, "review-card") or @role="article"][1]')
                                    rating_elem = parent.find_element(By.XPATH, './/span[@class="c-user-rating__rating"]')
                                    rating = rating_elem.text.strip()
                                    reviews[code] = rating
                                except:
                                    try:
                                        # 대체 방법
                                        rating_elem = elem.find_element(By.XPATH, './preceding::span[@class="c-user-rating__rating"][1]')
                                        rating = rating_elem.text.strip()
                                        reviews[code] = rating
                                    except:
                                        reviews[code] = ""
                        except:
                            continue
                    
                    # 다음 페이지
                    try:
                        next_page_btn = self.driver.find_element(
                            By.XPATH,
                            f'//button[@aria-label="Page {page_num + 1}"]'
                        )
                        next_page_btn.click()
                        time.sleep(2)
                        page_num += 1
                    except:
                        break
                        
                except Exception as e:
                    break
            
            print(f"  ✓ GG: {len(reviews)}개 리뷰 수집 완료")
            return reviews
            
        except Exception as e:
            print(f"  ✗ GG 수집 실패: {e}")
            import traceback
            traceback.print_exc()
            return reviews

        """KLOOK 필터 초기 설정 (한 번만 실행)"""
        try:
            from selenium.webdriver.common.keys import Keys
            date_str = tour_date.strftime("%Y-%m-%d")
            
            # 1. Product 필터에서 "Participation time" 선택
            try:
                print("  → Product 필터 열기...")
                product_dropdown = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[1]/form[2]/div[1]/div[2]/div/span'
                )
                product_dropdown.click()
                time.sleep(1)
                
                # "Participation time" 텍스트로 찾기 (ID가 동적이므로)
                print("  → Participation time 찾기...")
                participation_options = self.driver.find_elements(
                    By.XPATH,
                    '//li[contains(text(), "Participation time")]'
                )
                
                found = False
                for opt in participation_options:
                    if "Participation time" in opt.text:
                        print(f"  → 옵션 발견: {opt.text}")
                        opt.click()
                        time.sleep(0.5)
                        print("  ✓ Participation time 선택")
                        found = True
                        break
                
                if not found:
                    print("  ⚠ Participation time 못 찾음")
                    
            except Exception as e:
                print(f"  ✗ Product 필터 실패: {e}")
            
            # 2. 날짜 입력 (캘린더 팝업 안의 input 사용)
            try:
                print("  → 날짜 입력 시작...")
                
                # 메인 input 클릭해서 팝업 열기
                main_input = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[1]/form[2]/div[2]/div[2]/div/span/span/span/input[1]'
                )
                main_input.click()
                time.sleep(1)
                
                # 팝업 시작 날짜 input
                popup_start_input = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div/div/div[1]/div[1]/div[1]/div/input'))
                )
                popup_start_input.click()
                time.sleep(0.2)
                popup_start_input.send_keys(Keys.CONTROL + 'a')
                popup_start_input.send_keys(date_str)
                time.sleep(0.3)
                print(f"  ✓ 시작 날짜 입력: {date_str}")
                
                # 팝업 종료 날짜 input
                popup_end_input = self.driver.find_element(
                    By.XPATH,
                    '/html/body/div[3]/div/div/div/div/div[1]/div[2]/div[1]/div/input'
                )
                popup_end_input.click()
                time.sleep(0.2)
                popup_end_input.send_keys(Keys.CONTROL + 'a')
                popup_end_input.send_keys(date_str)
                time.sleep(0.3)
                print(f"  ✓ 종료 날짜 입력: {date_str}")
                
            except Exception as e:
                print(f"  ✗ 날짜 설정 실패: {e}")
            
            # 3. Search 버튼 클릭
            try:
                search_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[2]/button[1]'))
                )
                search_btn.click()
                time.sleep(3)
                print("  ✓ 검색 실행")
            except Exception as e:
                print(f"  ✗ Search 버튼 클릭 실패: {e}")
                return False
            
            # 50/page 설정은 포기 (ID가 계속 바뀜)
            print("  ✅ KLOOK 초기 설정 완료 (10개/페이지로 검색)")
            return True
                
        except Exception as e:
            print(f"  ✗ 필터 설정 실패: {e}")
            return False
        """KLOOK 필터 초기 설정 (한 번만 실행)"""
        try:
            from selenium.webdriver.common.keys import Keys
            date_str = tour_date.strftime("%Y-%m-%d")
            
            # 1. Product 필터에서 "Participation time" 선택
            try:
                print("  → Product 필터 열기...")
                product_dropdown = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[1]/form[2]/div[1]/div[2]/div/span'
                )
                product_dropdown.click()
                time.sleep(1)
                
                # "Participation time" 텍스트로 찾기 (ID가 동적이므로)
                print("  → Participation time 찾기...")
                participation_options = self.driver.find_elements(
                    By.XPATH,
                    '//li[contains(text(), "Participation time")]'
                )
                
                found = False
                for opt in participation_options:
                    if "Participation time" in opt.text:
                        print(f"  → 옵션 발견: {opt.text}")
                        opt.click()
                        time.sleep(0.5)
                        print("  ✓ Participation time 선택")
                        found = True
                        break
                
                if not found:
                    print("  ⚠ Participation time 못 찾음")
                    
            except Exception as e:
                print(f"  ✗ Product 필터 실패: {e}")
                import traceback
                traceback.print_exc()
            
            # 2. 날짜 입력 (캘린더 팝업 안의 input 사용)
            try:
                print("  → 날짜 입력 시작...")
                
                # 메인 input 클릭해서 팝업 열기
                print("  → 메인 input 찾기...")
                main_input = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[1]/form[2]/div[2]/div[2]/div/span/span/span/input[1]'
                )
                print("  → 메인 input 클릭...")
                main_input.click()
                time.sleep(1)
                print("  ✓ 캘린더 팝업 열림")
                
                # 팝업 시작 날짜 input
                print("  → 팝업 시작 날짜 input 찾기...")
                popup_start_input = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div/div/div[1]/div[1]/div[1]/div/input'))
                )
                print("  → 팝업 시작 날짜 input 발견")
                
                popup_start_input.click()
                print("  → 클릭 완료")
                time.sleep(0.2)
                
                # Ctrl+A로 전체 선택
                print("  → Ctrl+A 전송...")
                popup_start_input.send_keys(Keys.CONTROL + 'a')
                time.sleep(0.1)
                
                # 새 날짜 입력
                print(f"  → 날짜 입력 중: {date_str}")
                popup_start_input.send_keys(date_str)
                time.sleep(0.3)
                print(f"  ✓ 시작 날짜 입력 완료: {date_str}")
                
                # 팝업 종료 날짜 input
                print("  → 팝업 종료 날짜 input 찾기...")
                popup_end_input = self.driver.find_element(
                    By.XPATH,
                    '/html/body/div[3]/div/div/div/div/div[1]/div[2]/div[1]/div/input'
                )
                print("  → 팝업 종료 날짜 input 발견")
                
                popup_end_input.click()
                print("  → 클릭 완료")
                time.sleep(0.2)
                
                # Ctrl+A로 전체 선택
                print("  → Ctrl+A 전송...")
                popup_end_input.send_keys(Keys.CONTROL + 'a')
                time.sleep(0.1)
                
                # 새 날짜 입력
                print(f"  → 날짜 입력 중: {date_str}")
                popup_end_input.send_keys(date_str)
                time.sleep(0.3)
                print(f"  ✓ 종료 날짜 입력 완료: {date_str}")
                
            except Exception as e:
                print(f"  ✗ 날짜 설정 실패: {e}")
                import traceback
                traceback.print_exc()
            
            # 3. Search 버튼 클릭
            try:
                print("  → Search 버튼 클릭...")
                search_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[2]/button[1]'))
                )
                search_btn.click()
                time.sleep(3)
                print("  ✓ 검색 실행")
            except Exception as e:
                print(f"  ✗ Search 버튼 클릭 실패: {e}")
                return False
            
            
            print("  ✅ KLOOK 초기 설정 완료")
            return True
                
        except Exception as e:
            print(f"  ✗ 필터 설정 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def find_booking_in_klook_table(self, booking_code):
        """현재 페이지에서 예약번호 검색"""
        try:
            rows = self.driver.find_elements(
                By.XPATH,
                '//*[@id="klook-content"]/div/div[2]/div/div/div/div/div/div/div/div/div/div/table/tbody/tr'
            )
            
            for row in rows:
                try:
                    row_code = row.find_element(By.XPATH, './td[1]/a').text.strip()
                    if row_code == booking_code:
                        # 별점 확인 (6번째 열)
                        rating_text = row.find_element(By.XPATH, './td[6]').text.strip()
                        if rating_text and rating_text.isdigit():
                            rating = int(rating_text)
                            return True, str(rating)
                        else:
                            return True, ""
                except:
                    continue
            
            return False, ""
        except Exception as e:
            print(f"    ✗ 테이블 검색 오류: {e}")
            return False, ""
    
    def check_klook(self, booking_code, tour_date):
        """KLOOK 리뷰 체크"""
        try:
            print(f"\n[KLOOK] {booking_code} (날짜: {tour_date.strftime('%Y-%m-%d')})")
            
            # 매 예약마다 검색 결과 새로고침 (1페이지로 돌아감)
            if self.klook_setup_done:
                print("  → 검색 결과 새로고침 (1페이지로 이동)...")
                # Search 버튼 다시 클릭
                try:
                    search_btn = self.driver.find_element(
                        By.XPATH,
                        '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[2]/button[1]'
                    )
                    search_btn.click()
                    time.sleep(2)
                except:
                    print("  ⚠ Search 버튼 다시 클릭 실패")
            
            # 현재 페이지에서 검색
            found, rating = self.find_booking_in_klook_table(booking_code)
            
            if found:
                if rating and int(rating) >= 4:
                    print(f"  ✅ 리뷰 있음: {rating}점")
                    return "YES", rating
                else:
                    print(f"  ❌ 리뷰 없음")
                    return "NO", ""
            
            # 못 찾았으면 다음 페이지들 확인
            max_pages = 20  # 최대 20페이지까지 확인
            for page_num in range(max_pages):
                try:
                    # 다음 페이지 버튼 (class 기반 - 페이지 숫자 변경에도 작동)
                    next_btn = None
                    
                    # 방법 1: class로 찾기 (가장 안정적, 페이지 번호 상관없음)
                    try:
                        next_btn = self.driver.find_element(
                            By.XPATH,
                            '//li[contains(@class, "ant-pagination-next") and not(contains(@class, "ant-pagination-disabled"))]/a'
                        )
                        print(f"    ✓ 다음 페이지 버튼 발견 (class 방식)")
                    except:
                        pass
                    
                    # 방법 2: 아이콘으로 찾기
                    if not next_btn:
                        try:
                            next_btn = self.driver.find_element(
                                By.XPATH,
                                '//button[@aria-label="Next Page"] | //a[@aria-label="Next Page"]'
                            )
                            print(f"    ✓ 다음 페이지 버튼 발견 (aria-label 방식)")
                        except:
                            pass
                    
                    # 방법 3: 텍스트로 찾기 (최후의 수단)
                    if not next_btn:
                        try:
                            # ">" 또는 "Next" 텍스트 찾기
                            all_links = self.driver.find_elements(By.XPATH, '//li[contains(@class, "ant-pagination")]/a')
                            for link in all_links:
                                if '›' in link.text or '>' in link.text or 'next' in link.text.lower():
                                    parent = link.find_element(By.XPATH, '..')
                                    if 'disabled' not in parent.get_attribute('class'):
                                        next_btn = link
                                        print(f"    ✓ 다음 페이지 버튼 발견 (텍스트 방식)")
                                        break
                        except:
                            pass
                    
                    if not next_btn:
                        print(f"  ❌ 다음 페이지 버튼 없음")
                        return "NO", ""
                    
                    # disabled 체크 (이중 확인)
                    try:
                        parent_li = next_btn.find_element(By.XPATH, '..')
                        if 'ant-pagination-disabled' in parent_li.get_attribute('class'):
                            print(f"  ❌ 예약번호 없음 (마지막 페이지)")
                            return "NO", ""
                    except:
                        pass
                    
                    # 다음 페이지로
                    next_btn.click()
                    time.sleep(2)
                    print(f"    → 페이지 {page_num + 2} 확인 중...")
                    
                    # 현재 페이지에서 검색
                    found, rating = self.find_booking_in_klook_table(booking_code)
                    
                    if found:
                        if rating and int(rating) >= 4:
                            print(f"  ✅ 리뷰 있음: {rating}점")
                            return "YES", rating
                        else:
                            print(f"  ❌ 리뷰 없음")
                            return "NO", ""
                            
                except Exception as e:
                    print(f"    ⚠ 페이지 이동 실패: {e}")
                    break
            
            print(f"  ❌ 예약번호 없음 (최대 페이지 도달)")
            return "NO", ""
            
        except Exception as e:
            print(f"  ✗ KLOOK 오류: {e}")
            return "ERROR", ""
        """KLOOK 리뷰 체크"""
        try:
            print(f"\n[KLOOK] {booking_code} (날짜: {tour_date.strftime('%Y-%m-%d')})")
            
            # 현재 페이지에서 검색
            found, rating = self.find_booking_in_klook_table(booking_code)
            
            if found:
                if rating and int(rating) >= 4:
                    print(f"  ✅ 리뷰 있음: {rating}점")
                    return "YES", rating
                else:
                    print(f"  ❌ 리뷰 없음")
                    return "NO", ""
            
            # 못 찾았으면 다음 페이지들 확인
            max_pages = 10  # 최대 10페이지까지만 확인
            for page_num in range(max_pages):
                try:
                    # 다음 페이지 버튼
                    next_btn = self.driver.find_element(
                        By.XPATH,
                        '//*[@id="klook-content"]/div/div[2]/div/div/ul/li[3]/a'
                    )
                    
                    # 회색이면(disabled) 더 이상 페이지 없음
                    parent_li = next_btn.find_element(By.XPATH, '..')
                    if 'ant-pagination-disabled' in parent_li.get_attribute('class'):
                        print(f"  ❌ 예약번호 없음 (마지막 페이지)")
                        return "NO", ""
                    
                    # 다음 페이지로
                    next_btn.click()
                    time.sleep(2)
                    print(f"    → 페이지 {page_num + 2} 확인 중...")
                    
                    # 현재 페이지에서 검색
                    found, rating = self.find_booking_in_klook_table(booking_code)
                    
                    if found:
                        if rating and int(rating) >= 4:
                            print(f"  ✅ 리뷰 있음: {rating}점")
                            return "YES", rating
                        else:
                            print(f"  ❌ 리뷰 없음")
                            return "NO", ""
                            
                except Exception as e:
                    print(f"    ⚠ 페이지 이동 실패: {e}")
                    break
            
            print(f"  ❌ 예약번호 없음 (최대 페이지 도달)")
            return "NO", ""
            
        except Exception as e:
            print(f"  ✗ KLOOK 오류: {e}")
            return "ERROR", ""
    
    def check_kkday(self, booking_code, tour_date):
        """KKDAY 리뷰 체크"""
        try:
            print(f"\n[KKDAY] {booking_code}")
            
            # KKDAY 리뷰 페이지로 이동
            self.driver.get("https://scm.kkday.com/v1/en/comment/index")
            time.sleep(2)
            
            # 예약번호 입력
            try:
                order_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="orderMid"]'))
                )
                order_input.clear()
                order_input.send_keys(booking_code)
            except:
                print("  ✗ 입력란 찾기 실패")
                return "ERROR", ""
            
            # 검색 버튼 클릭
            try:
                search_btn = self.driver.find_element(By.XPATH, '//*[@id="searchBtn"]')
                search_btn.click()
                time.sleep(3)
            except:
                print("  ✗ 검색 버튼 클릭 실패")
                return "ERROR", ""
            
            # 결과 확인
            try:
                result_div = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="defaultLayout"]/div/section[2]/div[2]/div[2]/div/div/div[1]/div/div/div[1]/div[2]'))
                )
                result_text = result_div.text
                
                # "rating score:" 있는지 확인
                if "rating score:" in result_text.lower() or "Booking no.:" in result_text:
                    # 채워진 별만 세기 (fa-star, fa-star-o 제외)
                    filled_stars = result_div.find_elements(
                        By.XPATH,
                        './/p[1]/i[contains(@class, "fa-star") and not(contains(@class, "fa-star-o"))]'
                    )
                    star_count = len(filled_stars)
                    
                    if star_count > 0:
                        print(f"  ✅ 리뷰 있음: {star_count}점")
                        return "YES", str(star_count)
                
                print(f"  ❌ 리뷰 없음")
                return "NO", ""
                    
            except TimeoutException:
                print(f"  ❌ 검색 결과 없음")
                return "NO", ""
        
        except Exception as e:
            print(f"  ✗ KKDAY 오류: {e}")
            return "ERROR", ""
    
    def setup_gg_filters(self, tour_date):
        """GG 필터 초기 설정 (한 번만 실행)"""
        try:
            from datetime import timedelta
            
            # GG 리뷰 페이지로 이동
            self.driver.get("https://supplier.getyourguide.com/performance/reviews")
            time.sleep(3)
            
            # More Filters 클릭
            try:
                more_filters = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="__nuxt"]/div/div/main/div[1]/div/div[2]/div[1]/div/div[3]/button'))
                )
                more_filters.click()
                time.sleep(1)
                print("  ✓ More Filters 열림")
            except:
                print("  ⚠ More Filters 버튼 없음")
                return False
            
            # Activity date 선택 (전날 ~ 당일)
            try:
                # 전날 계산
                prev_day = tour_date - timedelta(days=1)
                prev_day_num = prev_day.day
                curr_day_num = tour_date.day
                
                print(f"  → 날짜 선택: {prev_day_num}일 ~ {curr_day_num}일")
                
                # Activity date 캘린더 열기
                calendar_btn = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="date-range"]/span/span/span'
                )
                calendar_btn.click()
                time.sleep(1)
                print("  ✓ 캘린더 열림")
                
                # 전날 선택 (시작일)
                prev_day_cell = self.driver.find_element(
                    By.XPATH,
                    f'//span[@class="p-datepicker-day" and text()="{prev_day_num}"]'
                )
                prev_day_cell.click()
                time.sleep(0.3)
                print(f"  ✓ {prev_day_num}일 선택")
                
                # 당일 선택 (종료일)
                curr_day_cell = self.driver.find_element(
                    By.XPATH,
                    f'//span[@class="p-datepicker-day" and text()="{curr_day_num}"]'
                )
                curr_day_cell.click()
                time.sleep(1)
                print(f"  ✓ {curr_day_num}일 선택 (범위 완료)")
                
                # 결과 로딩 대기 (중요!)
                print("  → 결과 로딩 대기 중...")
                time.sleep(5)
                
                print("  ✅ GG 초기 설정 완료")
                return True
                
            except Exception as e:
                print(f"  ⚠ 날짜 선택 실패: {e}")
                return False
                
        except Exception as e:
            print(f"  ✗ GG 필터 설정 실패: {e}")
            return False
    
    def find_booking_in_gg_page(self, booking_code):
        """현재 페이지에서 예약번호 검색"""
        try:
            # Show details 버튼들 찾기
            show_buttons = self.driver.find_elements(By.XPATH, '//button[contains(., "Show details")]')
            
            # 버튼이 있으면 클릭해서 열기
            if show_buttons:
                print(f"    → {len(show_buttons)}개 리뷰 확인 중...")
                for btn in show_buttons:
                    try:
                        # 버튼이 보이도록 스크롤
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                        time.sleep(0.2)
                        btn.click()
                        time.sleep(0.3)
                    except:
                        continue
            else:
                # 버튼 없으면 이미 열려있음
                print(f"    → 이미 열려있는 리뷰 검색 중...")
            
            # 예약번호 검색 (열려있든 닫혀있든)
            booking_elems = self.driver.find_elements(
                By.XPATH,
                '//a[contains(@href, "booking") or contains(text(), "GYG")]'
            )
            
            print(f"    → {len(booking_elems)}개 예약번호 발견")
            
            for elem in booking_elems:
                try:
                    found_code = elem.text.strip()
                    if found_code == booking_code:
                        print(f"    ✓ 예약번호 매칭: {found_code}")
                        # 별점 확인 - 같은 컨테이너 안에서 찾기
                        try:
                            # 방법 1: 가까운 부모 div에서 찾기
                            try:
                                parent = elem.find_element(By.XPATH, './ancestor::div[contains(@class, "c-review") or contains(@class, "review-card") or @role="article"][1]')
                                rating_elem = parent.find_element(By.XPATH, './/span[@class="c-user-rating__rating"]')
                                rating_text = rating_elem.text.strip()
                                print(f"    → 별점 발견 (방법1): {rating_text}점")
                            except:
                                # 방법 2: 같은 레벨에서 앞쪽에 있는 별점 찾기
                                rating_elem = elem.find_element(By.XPATH, './preceding::span[@class="c-user-rating__rating"][1]')
                                rating_text = rating_elem.text.strip()
                                print(f"    → 별점 발견 (방법2): {rating_text}점")
                            
                            if rating_text and rating_text.replace('.', '').isdigit():
                                rating = int(float(rating_text))
                                return True, str(rating)
                            else:
                                return True, rating_text
                                
                        except Exception as e:
                            print(f"    ⚠ 별점 추출 실패: {e}")
                            return True, ""
                except:
                    continue
            
            return False, ""
        except Exception as e:
            print(f"    ✗ 페이지 검색 오류: {e}")
            return False, ""
    
    def check_gg(self, booking_code, tour_date):
        """GetYourGuide 리뷰 체크"""
        try:
            print(f"\n[GG] {booking_code} (날짜: {tour_date.strftime('%Y-%m-%d')})")
            
            # 항상 1페이지로 이동
            if self.gg_setup_done:
                print("  → 1페이지로 이동...")
                try:
                    # 1페이지 버튼 클릭 (있으면)
                    page1_btn = self.driver.find_element(
                        By.XPATH,
                        '//button[@aria-label="Page 1"]'
                    )
                    page1_btn.click()
                    time.sleep(2)
                except:
                    # 1페이지 버튼 없으면 이미 1페이지
                    pass
                
                # 맨 위로 스크롤
                self.driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(1)
            
            # 현재 페이지에서 검색
            found, rating = self.find_booking_in_gg_page(booking_code)
            
            if found:
                if rating and rating.replace('.', '').isdigit() and int(float(rating)) >= 4:
                    print(f"  ✅ 리뷰 있음: {rating}점")
                    return "YES", rating
                elif rating:
                    print(f"  ✅ 리뷰 있음: {rating}점")
                    return "YES", rating
                else:
                    print(f"  ✅ 리뷰 있음 (별점 미확인)")
                    return "YES", ""
            
            # 못 찾았으면 다음 페이지들 확인
            max_pages = 10
            for page_num in range(2, max_pages + 1):
                try:
                    # 페이지 버튼 찾기 (숫자로)
                    page_btn = self.driver.find_element(
                        By.XPATH,
                        f'//button[@aria-label="Page {page_num}"]'
                    )
                    
                    page_btn.click()
                    time.sleep(2)
                    print(f"    → 페이지 {page_num} 확인 중...")
                    
                    # 현재 페이지에서 검색
                    found, rating = self.find_booking_in_gg_page(booking_code)
                    
                    if found:
                        if rating and rating.replace('.', '').isdigit() and int(float(rating)) >= 4:
                            print(f"  ✅ 리뷰 있음: {rating}점")
                            return "YES", rating
                        elif rating:
                            print(f"  ✅ 리뷰 있음: {rating}점")
                            return "YES", rating
                        else:
                            print(f"  ✅ 리뷰 있음 (별점 미확인)")
                            return "YES", ""
                            
                except:
                    # 더 이상 페이지 없음
                    break
            
            print(f"  ❌ 예약번호 없음")
            return "NO", ""
                
        except Exception as e:
            print(f"  ✗ GG 오류: {e}")
            import traceback
            traceback.print_exc()
            return "ERROR", ""
        """GetYourGuide 리뷰 체크"""
        try:
            print(f"\n[GG] {booking_code} (날짜: {tour_date.strftime('%Y-%m-%d')})")
            
            # GG 리뷰 페이지로 이동
            self.driver.get("https://supplier.getyourguide.com/performance/reviews")
            time.sleep(3)
            
            # More Filters 클릭
            try:
                more_filters = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="__nuxt"]/div/div/main/div[1]/div/div[2]/div[1]/div/div[3]/button'))
                )
                more_filters.click()
                time.sleep(1)
                print("  ✓ More Filters 열림")
            except:
                print("  ⚠ More Filters 버튼 없음")
            
            # Activity date 선택 (전날 ~ 당일)
            try:
                from datetime import timedelta
                
                # 전날 계산
                prev_day = tour_date - timedelta(days=1)
                prev_day_num = prev_day.day
                curr_day_num = tour_date.day
                
                print(f"  → 날짜 선택: {prev_day_num}일 ~ {curr_day_num}일")
                
                # Activity date 캘린더 열기
                calendar_btn = self.driver.find_element(
                    By.XPATH,
                    '//*[@id="date-range"]/span/span/span'
                )
                calendar_btn.click()
                time.sleep(1)
                print("  ✓ 캘린더 열림")
                
                # 전날 선택 (시작일)
                prev_day_cell = self.driver.find_element(
                    By.XPATH,
                    f'//span[@class="p-datepicker-day" and text()="{prev_day_num}"]'
                )
                prev_day_cell.click()
                time.sleep(0.3)
                print(f"  ✓ {prev_day_num}일 선택")
                
                # 당일 선택 (종료일)
                curr_day_cell = self.driver.find_element(
                    By.XPATH,
                    f'//span[@class="p-datepicker-day" and text()="{curr_day_num}"]'
                )
                curr_day_cell.click()
                time.sleep(0.5)
                print(f"  ✓ {curr_day_num}일 선택 (범위 완료)")
                
            except Exception as e:
                print(f"  ⚠ 날짜 선택 실패: {e}")
                print("  → 전체 날짜로 검색 진행")
            
            time.sleep(2)
            
            # Show details 버튼들 찾기
            try:
                # 페이지의 모든 Show details 버튼 찾기
                show_buttons = self.driver.find_elements(By.XPATH, '//button//span[contains(text(), "Show details")]/..')
                
                for btn_idx, btn in enumerate(show_buttons[:10]):  # 최대 10개 확인
                    try:
                        btn.click()
                        time.sleep(1)
                        
                        # 예약번호 확인
                        booking_elem = self.driver.find_element(By.XPATH, f'//*[@id="__nuxt"]/div/div/main/div[1]/div/div[2]/div[2]/div/div/div[{btn_idx+1}]/ul/li[2]/div[2]/a')
                        found_code = booking_elem.text.strip()
                        
                        if found_code == booking_code:
                            # 별점 확인 (c-user-rating__rating 클래스 사용)
                            try:
                                rating_elem = self.driver.find_element(
                                    By.XPATH,
                                    f'//*[@id="__nuxt"]/div/div/main/div[1]/div/div[2]/div[2]/div/div/div[{btn_idx+1}]//span[@class="c-user-rating__rating"]'
                                )
                                rating_text = rating_elem.text.strip()
                                
                                if rating_text and rating_text.replace('.', '').isdigit():
                                    rating = int(float(rating_text))
                                    print(f"  ✅ 리뷰 있음: {rating}점")
                                    return "YES", str(rating)
                                else:
                                    print(f"  ✅ 리뷰 있음: {rating_text}점")
                                    return "YES", rating_text
                            except:
                                # 별점 못 찾으면 그냥 리뷰 있음으로 처리
                                print(f"  ✅ 리뷰 있음 (별점 미확인)")
                                return "YES", ""
                    except:
                        continue
                
                print(f"  ❌ 예약번호 없음")
                return "NO", ""
                
            except Exception as e:
                print(f"  ✗ Show details 처리 실패: {e}")
                return "ERROR", ""
        
        except Exception as e:
            print(f"  ✗ GG 오류: {e}")
            return "ERROR", ""
    
    def copy_results(self):
        """조회 결과를 클립보드에 복사"""
        try:
            # 텍스트 위젯에서 모든 내용 가져오기
            result_text = self.result_text.get(1.0, "end-1c")
            
            if not result_text.strip():
                messagebox.showwarning("경고", "복사할 결과가 없습니다.\n먼저 리뷰 조회를 완료하세요.")
                return
            
            # 클립보드에 복사
            self.root.clipboard_clear()
            self.root.clipboard_append(result_text)
            self.root.update()  # 클립보드 업데이트
            
            messagebox.showinfo("성공", "✅ 조회 결과가 클립보드에 복사되었습니다!\n\n다른 곳에 Ctrl+V로 붙여넣기 하세요.")
            
        except Exception as e:
            messagebox.showerror("오류", f"복사 실패:\n{e}")
    
    def quit_app(self):
        """프로그램 종료"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        self.root.quit()
        self.root.destroy()
    
    def run(self):
        """GUI 실행"""
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app)
        self.root.mainloop()


if __name__ == "__main__":
    print("=" * 60)
    print("Review Checker 시작")
    print("=" * 60)
    print("\n⚠️  먼저 크롬을 디버그 모드로 실행하세요:")
    print("\nWindows:")
    print('  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222')
    print("\nMac:")
    print('  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222')
    print("\n그 다음:")
    print("  1. KLOOK 로그인: https://merchant.klook.com/reviews")
    print("  2. KKDAY 로그인: https://scm.kkday.com/v1/en/comment/index")
    print("  3. GG 로그인: https://supplier.getyourguide.com/performance/reviews")
    print("=" * 60)
    print()
    
    app = ReviewCheckerGUI()
    app.run()
