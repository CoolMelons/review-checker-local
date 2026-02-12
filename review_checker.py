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
        self.root.geometry("700x1200")

        self.driver = None
        self.df = None

        # 노쇼 관련
        self.noshow_codes = set()          # No Show 탭에서 O인 Agency Code들
        self.noshow_teams = 0
        self.noshow_people = 0

        self.guide_groups = []
        self.guide_checkboxes = {}
        self.select_all_var = BooleanVar(value=True)

        self.klook_setup_done = False
        self.klook_current_date = None
        self.gg_setup_done = False
        self.gg_current_date = None

        self.setup_ui()

    def setup_ui(self):
        Label(self.root, text="📋 Review Checker", font=("Arial", 18, "bold")).pack(pady=15)

        frame1 = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        frame1.pack(fill="x", padx=20, pady=5)

        Label(frame1, text="1️⃣ 크롬 연결 (디버그 모드)", font=("Arial", 12, "bold")).pack(anchor="w")
        Label(frame1, text="⚠️ L, KK, GG 로그인 필요", font=("Arial", 9), fg="red").pack(anchor="w")

        self.chrome_status = StringVar(value="🔴 크롬 미연결")
        Label(frame1, textvariable=self.chrome_status, font=("Arial", 10)).pack(anchor="w", pady=5)

        Button(frame1, text="🔌 크롬 연결", command=self.connect_chrome,
               width=20, height=1, bg="#4CAF50", fg="white").pack(anchor="w")

        frame2 = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        frame2.pack(fill="x", padx=20, pady=5)

        Label(frame2, text="2️⃣ 엑셀 파일 선택 (Excel for Guides)", font=("Arial", 12, "bold")).pack(anchor="w")

        self.file_status = StringVar(value="📁 파일 미선택")
        Label(frame2, textvariable=self.file_status, font=("Arial", 10)).pack(anchor="w", pady=5)

        Button(frame2, text="📁 파일 선택", command=self.select_file,
               width=20, height=1, bg="#2196F3", fg="white").pack(anchor="w")

        self.guide_frame = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        self.guide_frame.pack(fill="both", expand=True, padx=20, pady=5)

        Label(self.guide_frame, text="조회할 가이드 선택:", font=("Arial", 12, "bold")).pack(anchor="w")

        self.select_all_check = Checkbutton(
            self.guide_frame,
            text="☑ 전체 선택",
            variable=self.select_all_var,
            command=self.toggle_all
        )
        self.select_all_check.pack(anchor="w", pady=5)

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

        Button(self.root, text="▶️ 선택한 가이드만 조회 시작",
               command=self.start_processing,
               width=30, height=2,
               bg="#FF9800", fg="white",
               font=("Arial", 11, "bold")).pack(pady=10)

        result_frame = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        result_frame.pack(fill="both", expand=True, padx=20, pady=5)

        Label(result_frame, text="📊 조회 결과", font=("Arial", 12, "bold")).pack(anchor="w")

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

        self.progress_var = StringVar(value="")
        Label(self.root, textvariable=self.progress_var, font=("Arial", 9)).pack(pady=5)

        button_frame = Frame(self.root)
        button_frame.pack(pady=5)

        Button(button_frame, text="📋 Copy",
               command=self.copy_results, width=20,
               bg="#9C27B0", fg="white").pack(side="left", padx=5)

        Button(button_frame, text="End",
               command=self.quit_app, width=15).pack(side="left", padx=5)

    def connect_chrome(self):
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

    # =========================
    # No Show 탭 처리 핵심
    # =========================
    def load_excel_with_noshow(self, file_path: str):
        """
        엑셀 전체 시트를 읽고,
        - 본 데이터 시트(df_main)
        - No Show 시트에서 'O'인 Agency Code들을 noshow_codes로 추출
        """
        xls = pd.read_excel(file_path, sheet_name=None)

        # 1) No Show 시트 찾기 (대소문자/공백 차이 대응)
        noshow_sheet_name = None
        for name in xls.keys():
            if str(name).strip().lower() in ["no show", "noshow", "no_show", "no-show"]:
                noshow_sheet_name = name
                break
            # 한국어 등 변형 대응
            if "no show" in str(name).strip().lower():
                noshow_sheet_name = name
                break

        # 2) 메인 시트 선택: No Show 제외하고 첫 번째를 메인으로
        main_sheet_name = None
        for name in xls.keys():
            if name == noshow_sheet_name:
                continue
            main_sheet_name = name
            break

        if main_sheet_name is None:
            raise ValueError("메인 데이터 시트를 찾을 수 없습니다. (No Show만 있는지 확인)")

        df_main = xls[main_sheet_name].copy()
        df_main = self.normalize_columns(df_main)

        # 3) No Show codes 추출
        noshow_codes = set()
        if noshow_sheet_name is not None:
            df_ns = xls[noshow_sheet_name].copy()
            df_ns.columns = [str(c).strip() for c in df_ns.columns]

            # 코드 컬럼 후보
            code_col = None
            for c in df_ns.columns:
                lc = c.lower()
                if lc in ["agency code", "booking code", "booking", "order", "order id", "reservation", "reservation code"]:
                    code_col = c
                    break
                if "code" in lc and code_col is None:
                    code_col = c

            # O/X 컬럼 후보
            flag_col = None
            for c in df_ns.columns:
                lc = c.lower().replace(" ", "")
                if lc in ["noshow", "no_show", "no-show"]:
                    flag_col = c
                    break
                # 그냥 첫 번째/두 번째 형태로 오는 경우도 있어서 텍스트 힌트
                if "show" in lc and flag_col is None:
                    flag_col = c

            # 만약 O/X가 특정 컬럼 없이 코드만 O인 행으로 들어오는 경우도 대비
            # -> flag_col이 없으면, df_ns 전체에서 'O'가 있는 행의 code_col을 잡는 방식
            if code_col is not None:
                if flag_col is not None:
                    for _, r in df_ns.iterrows():
                        code = str(r.get(code_col, "")).strip()
                        flag = str(r.get(flag_col, "")).strip().upper()
                        if not code:
                            continue
                        # 'O'만 노쇼로 인정 (X는 무시)
                        if flag == "O":
                            noshow_codes.add(code)
                else:
                    # flag_col이 없으면, 행 전체에 'O'가 포함된 경우를 노쇼로 보고 code_col 추출
                    for _, r in df_ns.iterrows():
                        code = str(r.get(code_col, "")).strip()
                        if not code:
                            continue
                        row_text = " ".join([str(v) for v in r.values]).upper()
                        if " O " in f" {row_text} " or row_text.strip() == "O":
                            noshow_codes.add(code)

        return df_main, noshow_codes, main_sheet_name, noshow_sheet_name

    def select_file(self):
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
            df, noshow_codes, main_sheet, noshow_sheet = self.load_excel_with_noshow(file_path)

            # 서울 필터링
            df = df[df["Area"].astype(str).str.strip().str.lower() == "seoul"].copy()

            # 데이터 준비
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Agency"] = df["Agency"].astype(str).str.strip()
            df["Agency Code"] = df["Agency Code"].astype(str).str.strip()
            if "People" in df.columns:
                df["People"] = pd.to_numeric(df["People"], errors="coerce").fillna(0).astype(int)

            # No Show 적용 (O인 코드 제외)
            self.noshow_codes = set([str(c).strip() for c in noshow_codes if str(c).strip()])
            if self.noshow_codes:
                df["__NOSHOW__"] = df["Agency Code"].astype(str).str.strip().isin(self.noshow_codes)
                self.noshow_teams = int(df["__NOSHOW__"].sum())
                self.noshow_people = int(df.loc[df["__NOSHOW__"], "People"].sum()) if "People" in df.columns else 0

                df = df[~df["__NOSHOW__"]].copy()
                df.drop(columns=["__NOSHOW__"], inplace=True, errors="ignore")
            else:
                self.noshow_teams = 0
                self.noshow_people = 0

            self.df = df

            ns_msg = ""
            if noshow_sheet is not None:
                ns_msg = f" | No Show(O) 제외: {self.noshow_teams}팀 {self.noshow_people}명"
            self.file_status.set(f"✅ 파일 로드 완료: {len(df)}개 예약{ns_msg}")

            self.extract_and_display_guides()

        except Exception as e:
            messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")

    def extract_and_display_guides(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.guide_groups = []
        self.guide_checkboxes = {}

        grouped = self.df.groupby(['Date', 'Product', 'Main Guide'])

        for (date_val, product, guide), group in grouped:
            self.guide_groups.append((date_val, product, guide))

            var = BooleanVar(value=True)
            self.guide_checkboxes[(date_val, product, guide)] = var

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

        messagebox.showinfo("완료", f"{len(self.guide_groups)}개 가이드 그룹을 찾았습니다.\n(No Show O는 자동 제외됨)")

    def display_results(self, stats):
        self.result_text.delete(1.0, "end")

        result = []
        result.append("=" * 60)
        result.append("📈 전체 통계")
        result.append("=" * 60)

        # 노쇼 제외 정보 표시
        if stats.get("noshow_total", 0) > 0:
            result.append(f"🚫 No Show(O) 제외: {stats['noshow_total']}팀 {stats['noshow_people']}명")

        result.append(f"👥 (노쇼 제외 후) 총 예약: {stats['total_teams']}팀 {stats['total_people']}명")

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

                for agency_code in ['L', 'KK', 'GG']:
                    agency_stat = guide_stat['agencies'][agency_code]
                    if agency_stat['total'] > 0:
                        agency_pct = (agency_stat['checked'] / agency_stat['total']) * 100
                        agency_avg = sum(agency_stat['ratings']) / len(agency_stat['ratings']) if agency_stat['ratings'] else 0
                        line = f"    └ {agency_code:15} {agency_stat['checked']:2}팀 / {agency_stat['total']:2}팀 ({agency_pct:5.1f}%)"
                        if agency_avg > 0:
                            line += f" - 평균 {agency_avg:.1f}점"
                        result.append(line)

                for other_agency, bookings in guide_stat['other_agencies'].items():
                    if len(bookings) > 0:
                        total_people = sum(b['people'] for b in bookings)
                        result.append(f"    └ {other_agency:15} {len(bookings):2}팀 / {total_people:3}명 (검색 필요)")

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

        if stats['other_agencies']:
            result.append("\n[개별 조회 필요 에이전시]")
            result.append("-" * 60)
            for agency_code, agency_data in stats['other_agencies'].items():
                result.append(f"  {agency_code:15} {agency_data['total']:2}팀")
                for booking in agency_data['bookings']:
                    result.append(f"    · {booking['code']} ({booking['guide']})")

        result.append("\n" + "=" * 60)
        self.result_text.insert("end", "\n".join(result))

    def toggle_all(self):
        select_all = self.select_all_var.get()
        for var in self.guide_checkboxes.values():
            var.set(select_all)

    def start_processing(self):
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!")
            return

        if self.df is None:
            messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
            return

        selected_guides = [key for key, var in self.guide_checkboxes.items() if var.get()]

        if not selected_guides:
            messagebox.showerror("오류", "최소 1개 이상의 가이드를 선택하세요!")
            return

        filtered_df = pd.DataFrame()
        for date_val, product, guide in selected_guides:
            mask = (
                (self.df['Date'] == date_val) &
                (self.df['Product'] == product) &
                (self.df['Main Guide'] == guide)
            )
            filtered_df = pd.concat([filtered_df, self.df[mask]])

        self.select_file_and_start(filtered_df)

    def select_file_and_start(self, df=None):
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!")
            return

        # 레거시: df가 None이면 여기서 파일 선택
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
                df, noshow_codes, main_sheet, noshow_sheet = self.load_excel_with_noshow(file_path)

                df = df[df["Area"].astype(str).str.strip().str.lower() == "seoul"].copy()

                df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
                df["Agency"] = df["Agency"].astype(str).str.strip()
                df["Agency Code"] = df["Agency Code"].astype(str).str.strip()
                if "People" in df.columns:
                    df["People"] = pd.to_numeric(df["People"], errors="coerce").fillna(0).astype(int)

                # No Show O 제외
                self.noshow_codes = set([str(c).strip() for c in noshow_codes if str(c).strip()])
                if self.noshow_codes:
                    df["__NOSHOW__"] = df["Agency Code"].astype(str).str.strip().isin(self.noshow_codes)
                    self.noshow_teams = int(df["__NOSHOW__"].sum())
                    self.noshow_people = int(df.loc[df["__NOSHOW__"], "People"].sum()) if "People" in df.columns else 0
                    df = df[~df["__NOSHOW__"]].copy()
                    df.drop(columns=["__NOSHOW__"], inplace=True, errors="ignore")
                else:
                    self.noshow_teams = 0
                    self.noshow_people = 0

            except Exception as e:
                messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")
                return

        try:
            df["Review_Status"] = ""
            df["Rating"] = ""
            df["Check"] = ""

            self.klook_setup_done = False
            self.klook_current_date = None
            self.gg_setup_done = False
            self.gg_current_date = None

            stats = {
                'noshow_total': self.noshow_teams,
                'noshow_people': self.noshow_people,

                'total_teams': 0,
                'total_people': 0,
                'total_checked': 0,
                'total_ratings': [],
                'agencies': {
                    'L': {'name': 'KLOOK', 'total': 0, 'checked': 0, 'ratings': []},
                    'KK': {'name': 'KKDAY', 'total': 0, 'checked': 0, 'ratings': []},
                    'GG': {'name': 'GetYourGuide', 'total': 0, 'checked': 0, 'ratings': []}
                },
                'guides': {},
                'other_agencies': {},
                'reviewed_total': 0,
                'reviewed_people': 0
            }

            progress_window = self.create_progress_window()
            progress_bar = progress_window.progress_bar
            progress_label = progress_window.label

            print("\n" + "=" * 80)
            print("📊 리뷰 조회 시작".center(80))
            print("=" * 80 + "\n")

            unique_dates = df['Date'].unique()
            all_reviews = {'L': {}, 'KK': {}, 'GG': {}}

            print("=" * 80)
            print("1단계: 날짜별 리뷰 수집")
            print("=" * 80)

            for date_val in unique_dates:
                print(f"\n📅 {pd.to_datetime(date_val).strftime('%Y-%m-%d')}")
                print("-" * 60)

                klook_reviews = self.collect_klook_reviews(date_val)
                all_reviews['L'][date_val] = klook_reviews

                all_reviews['KK'][date_val] = {}

                gg_reviews = self.collect_gg_reviews(date_val)
                all_reviews['GG'][date_val] = gg_reviews

            print("\n" + "=" * 80)
            print("2단계: 예약번호 매칭 및 출력")
            print("=" * 80)

            grouped = df.groupby(['Date', 'Product', 'Main Guide'])
            processed_count = 0
            total = len(df)

            current_date = None

            for (date_val, product, guide), group in grouped:
                if current_date != date_val:
                    if current_date is not None:
                        print()
                    print(f"\n{'=' * 80}")
                    print(f"📅 {date_val.strftime('%Y-%m-%d (%A)')}")
                    print(f"{'=' * 80}\n")
                    current_date = date_val

                people_count = group['People'].sum() if 'People' in group.columns else 0
                team_count = len(group)

                print(f"📍 투어: {product}")
                print(f"👤 가이드: {guide}")
                print(f"👥 총: {team_count}팀 {people_count}명\n")

                stats['total_teams'] += team_count
                stats['total_people'] += people_count

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
                        'other_agencies': {}
                    }
                stats['guides'][guide]['total'] += team_count

                for agency in ['L', 'KK', 'GG']:
                    agency_group = group[group['Agency'] == agency]
                    if len(agency_group) == 0:
                        continue

                    print(f"[{agency}]")
                    print("-" * 60)

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

                        status = "NO"
                        rating = ""

                        if agency == "L" or agency == "GG":
                            date_reviews = all_reviews[agency].get(date, {})
                            if code in date_reviews:
                                status = "YES"
                                rating = date_reviews[code]
                        elif agency == "KK":
                            status, rating = self.check_kkday(code, date)
                        else:
                            status = "SKIP"

                        df.at[idx, "Review_Status"] = status
                        df.at[idx, "Rating"] = rating

                        stats['guides'][guide]['agencies'][agency]['total'] += 1
                        stats['agencies'][agency]['total'] += 1
                        stats['reviewed_total'] += 1
                        stats['reviewed_people'] += people

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

                    current_total = len(agency_group)
                    if current_total > 0:
                        pct = (current_checked / current_total) * 100
                        avg = sum(current_ratings) / len(current_ratings) if current_ratings else 0
                        print(f"\n  📊 {current_checked}/{current_total}팀 ({pct:.1f}%)", end="")
                        if avg > 0:
                            print(f" - 평균 {avg:.1f}점\n")
                        else:
                            print("\n")

                other_group = group[~group['Agency'].isin(['L', 'KK', 'GG'])]
                for idx, row in other_group.iterrows():
                    agency = row["Agency"]
                    code = row["Agency Code"]
                    people = row.get("People", 0)

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

                    if agency not in stats['guides'][guide]['other_agencies']:
                        stats['guides'][guide]['other_agencies'][agency] = []

                    stats['guides'][guide]['other_agencies'][agency].append({
                        'code': code,
                        'people': people
                    })

            progress_window.window.destroy()
            self.display_results(stats)

            if stats['reviewed_total'] > 0:
                if stats['total_ratings']:
                    final_msg = (
                        f"✅ 완료!\n\n"
                        f"(No Show O 제외)\n"
                        f"리뷰 확인: {stats['total_checked']}/{stats['reviewed_total']}팀 ({stats['total_checked']/stats['reviewed_total']*100:.1f}%)\n"
                        f"평균 별점: {sum(stats['total_ratings'])/len(stats['total_ratings']):.1f}점"
                    )
                else:
                    final_msg = (
                        f"✅ 완료!\n\n"
                        f"(No Show O 제외)\n"
                        f"리뷰 확인: {stats['total_checked']}/{stats['reviewed_total']}팀"
                    )
            else:
                final_msg = "✅ 완료!\n\n(No Show O 제외)\n리뷰 조회 대상(L/KK/GG)이 없습니다."

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
        window = Toplevel(self.root)
        window.title("처리 중...")
        window.geometry("400x100")

        label = Label(window, text="시작 중...", font=("Arial", 10))
        label.pack(pady=10)

        progress_bar = Progressbar(window, length=350, mode="determinate")
        progress_bar.pack(pady=10)

        window.protocol("WM_DELETE_WINDOW", lambda: None)

        window.progress_bar = progress_bar
        window.label = label
        window.window = window

        return window

    def normalize_columns(self, df):
        df.columns = [str(c).strip() for c in df.columns]
        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(f"필수 컬럼 누락: {missing}")
        return df

    def collect_klook_reviews(self, date):
        reviews = {}
        try:
            print(f"\n🔍 KLOOK 리뷰 수집 중... (날짜: {date.strftime('%Y-%m-%d')})")

            self.driver.get("https://merchant.klook.com/reviews")
            time.sleep(2)

            try:
                date_str = date.strftime("%Y-%m-%d")

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

                search_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="klook-content"]/div/div[1]/div[1]/div/div[2]/button[1]'))
                )
                search_btn.click()
                time.sleep(3)

            except Exception as e:
                print(f"  ⚠ 날짜 필터 설정 실패: {e}")

            page_num = 1
            while page_num <= 20:
                try:
                    rows = self.driver.find_elements(
                        By.XPATH,
                        '//*[@id="klook-content"]/div/div[2]/div/div/div/div/div/div/div/div/div/div/table/tbody/tr'
                    )

                    for row in rows:
                        try:
                            code = row.find_element(By.XPATH, './td[1]/a').text.strip()
                            rating_text = row.find_element(By.XPATH, './td[6]').text.strip()

                            if code:
                                reviews[code] = rating_text if rating_text.isdigit() else ""
                        except:
                            continue

                    print(f"  → 페이지 {page_num}: {len(rows)}개 리뷰")

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

                except Exception:
                    break

            print(f"  ✓ KLOOK: {len(reviews)}개 리뷰 수집 완료")
            return reviews

        except Exception as e:
            print(f"  ✗ KLOOK 수집 실패: {e}")
            import traceback
            traceback.print_exc()
            return reviews

    def collect_kkday_reviews(self, date):
        reviews = {}
        try:
            print(f"\n🔍 KKDAY 리뷰 수집 중... (날짜: {date.strftime('%Y-%m-%d')})")
            print(f"  ⚠ KKDAY는 개별 조회 방식 유지")
            return reviews
        except Exception as e:
            print(f"  ✗ KKDAY 수집 실패: {e}")
            return reviews

    def collect_gg_reviews(self, date):
        reviews = {}
        try:
            from datetime import timedelta

            print(f"\n🔍 GG 리뷰 수집 중... (날짜: {date.strftime('%Y-%m-%d')})")

            self.driver.get("https://supplier.getyourguide.com/performance/reviews")
            time.sleep(3)

            try:
                more_filters = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="__nuxt"]/div/div/main/div[1]/div/div[2]/div[1]/div/div[3]/button'))
                )
                more_filters.click()
                time.sleep(1)
            except:
                pass

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
                time.sleep(5)

            except Exception as e:
                print(f"  ⚠ 날짜 선택 실패: {e}")

            page_num = 1
            while page_num <= 10:
                try:
                    show_buttons = self.driver.find_elements(By.XPATH, '//button[contains(., "Show details")]')
                    for btn in show_buttons:
                        try:
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                            time.sleep(0.2)
                            btn.click()
                            time.sleep(0.3)
                        except:
                            continue

                    booking_elems = self.driver.find_elements(
                        By.XPATH,
                        '//a[contains(@href, "booking") or contains(text(), "GYG")]'
                    )

                    print(f"  → 페이지 {page_num}: {len(booking_elems)}개 예약번호 발견")

                    for elem in booking_elems:
                        try:
                            code = elem.text.strip()
                            if code.startswith("GYG"):
                                try:
                                    parent = elem.find_element(By.XPATH, './ancestor::div[contains(@class, "c-review") or contains(@class, "review-card") or @role="article"][1]')
                                    rating_elem = parent.find_element(By.XPATH, './/span[@class="c-user-rating__rating"]')
                                    rating = rating_elem.text.strip()
                                    reviews[code] = rating
                                except:
                                    try:
                                        rating_elem = elem.find_element(By.XPATH, './preceding::span[@class="c-user-rating__rating"][1]')
                                        rating = rating_elem.text.strip()
                                        reviews[code] = rating
                                    except:
                                        reviews[code] = ""
                        except:
                            continue

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

                except Exception:
                    break

            print(f"  ✓ GG: {len(reviews)}개 리뷰 수집 완료")
            return reviews

        except Exception as e:
            print(f"  ✗ GG 수집 실패: {e}")
            import traceback
            traceback.print_exc()
            return reviews

    def check_kkday(self, booking_code, tour_date):
        try:
            print(f"\n[KKDAY] {booking_code}")

            self.driver.get("https://scm.kkday.com/v1/en/comment/index")
            time.sleep(2)

            try:
                order_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="orderMid"]'))
                )
                order_input.clear()
                order_input.send_keys(booking_code)
            except:
                print("  ✗ 입력란 찾기 실패")
                return "ERROR", ""

            try:
                search_btn = self.driver.find_element(By.XPATH, '//*[@id="searchBtn"]')
                search_btn.click()
                time.sleep(3)
            except:
                print("  ✗ 검색 버튼 클릭 실패")
                return "ERROR", ""

            try:
                result_div = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="defaultLayout"]/div/section[2]/div[2]/div[2]/div/div/div[1]/div/div/div[1]/div[2]'))
                )
                result_text = result_div.text

                if "rating score:" in result_text.lower() or "Booking no.:" in result_text:
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

    def copy_results(self):
        try:
            result_text = self.result_text.get(1.0, "end-1c")

            if not result_text.strip():
                messagebox.showwarning("경고", "복사할 결과가 없습니다.\n먼저 리뷰 조회를 완료하세요.")
                return

            self.root.clipboard_clear()
            self.root.clipboard_append(result_text)
            self.root.update()

            messagebox.showinfo("성공", "✅ 조회 결과가 클립보드에 복사되었습니다!\n\n다른 곳에 Ctrl+V로 붙여넣기 하세요.")

        except Exception as e:
            messagebox.showerror("오류", f"복사 실패:\n{e}")

    def quit_app(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        self.root.quit()
        self.root.destroy()

    def run(self):
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
