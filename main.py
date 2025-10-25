import openpyxl
from openpyxl import Workbook
from faker import Faker
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime, timedelta
import os

class DummyDataGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("ダミーデータ作成")
        self.root.geometry("1000x750")
        
        # 背景色を設定
        self.root.configure(bg='#f0f4f8')
        
        # 日本語と英語のFakerを初期化
        self.fake_ja = Faker('ja_JP')
        self.fake_en = Faker('en_US')
        
        # 現在の年度を計算（4月始まり）
        today = datetime.now()
        if today.month >= 4:
            self.current_year = today.year
        else:
            self.current_year = today.year - 1
        
        # スタイルを設定
        self.setup_style()
        self.create_widgets()
    
    def setup_style(self):
        """カスタムスタイルを設定"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # LabelFrameのスタイル
        style.configure('Custom.TLabelframe', 
                       background='#ffffff',
                       borderwidth=2,
                       relief='solid')
        style.configure('Custom.TLabelframe.Label', 
                       background='#ffffff',
                       foreground='#2c3e50',
                       font=('Arial', 11, 'bold'))
        
        # Checkbuttonのスタイル - 大きくする
        style.configure('Custom.TCheckbutton',
                       background='#ffffff',
                       foreground='#2c3e50',
                       font=('Arial', 11))
    
    def create_widgets(self):
        # メインコンテナ
        container = tk.Frame(self.root, bg='#f0f4f8')
        container.pack(fill="both", expand=True)
        
        # キャンバスとスクロールバーを作成
        canvas = tk.Canvas(container, bg='#f0f4f8', highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#f0f4f8')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # マウスホイールでスクロール
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # ファイル選択
        file_frame = ttk.LabelFrame(scrollable_frame, text="Excelファイル", 
                                    padding=15, style='Custom.TLabelframe')
        file_frame.pack(fill="x", padx=15, pady=8)
        
        self.file_path = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=80, font=('Arial', 10))
        file_entry.pack(side="left", padx=5)
        
        browse_btn = tk.Button(file_frame, text="参照", command=self.select_file,
                              bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                              padx=15, pady=5, relief='flat', cursor='hand2',
                              activebackground='#2980b9', activeforeground='white')
        browse_btn.pack(side="left", padx=5)
        
        # シート名と共通設定を横並び
        settings_container = tk.Frame(scrollable_frame, bg='#f0f4f8')
        settings_container.pack(fill="x", padx=15, pady=8)
        
        # シート名
        sheet_frame = ttk.LabelFrame(settings_container, text="シート名", 
                                     padding=15, style='Custom.TLabelframe')
        sheet_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        self.sheet_name = tk.StringVar(value="Sheet1")
        ttk.Entry(sheet_frame, textvariable=self.sheet_name, width=20, font=('Arial', 10)).pack()
        
        # 共通設定
        common_frame = ttk.LabelFrame(settings_container, text="共通設定", 
                                      padding=15, style='Custom.TLabelframe')
        common_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        
        common_inner = tk.Frame(common_frame, bg='#ffffff')
        common_inner.pack()
        
        tk.Label(common_inner, text="開始行:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=0, sticky="w", padx=5)
        self.start_row = tk.StringVar(value="2")
        ttk.Entry(common_inner, textvariable=self.start_row, width=10, 
                 font=('Arial', 10)).grid(row=0, column=1, padx=5)
        
        tk.Label(common_inner, text="件数:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=2, sticky="w", padx=20)
        self.data_count = tk.StringVar(value="100")
        ttk.Entry(common_inner, textvariable=self.data_count, width=10, 
                 font=('Arial', 10)).grid(row=0, column=3, padx=5)
        
        # データ項目を2段レイアウトに
        # 左列
        left_column = tk.Frame(scrollable_frame, bg='#f0f4f8')
        left_column.pack(side="left", fill="both", expand=True, padx=(15, 7))
        
        # 右列
        right_column = tk.Frame(scrollable_frame, bg='#f0f4f8')
        right_column.pack(side="left", fill="both", expand=True, padx=(7, 15))
        
        # 学生番号設定（左列）
        self.student_id_enabled = tk.BooleanVar(value=False)
        student_id_frame = ttk.LabelFrame(left_column, 
                                          text="学生番号（2から始まる7桁の数字）", 
                                          padding=15, style='Custom.TLabelframe')
        student_id_frame.pack(fill="x", pady=8)
        
        student_id_inner = tk.Frame(student_id_frame, bg='#ffffff')
        student_id_inner.pack()
        
        student_id_check = tk.Checkbutton(student_id_inner, text="生成する", 
                                         variable=self.student_id_enabled,
                                         bg='#ffffff', fg='#2c3e50',
                                         font=('Arial', 11), selectcolor='#ffffff',
                                         activebackground='#ffffff',
                                         activeforeground='#3498db',
                                         cursor='hand2', padx=5, pady=5)
        student_id_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(student_id_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.student_id_column = tk.StringVar(value="A")
        ttk.Entry(student_id_inner, textvariable=self.student_id_column, 
                 width=10, font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # 名前設定（左列）
        self.name_enabled = tk.BooleanVar(value=False)
        name_frame = ttk.LabelFrame(left_column, text="名前", 
                                    padding=15, style='Custom.TLabelframe')
        name_frame.pack(fill="x", pady=8)
        
        name_inner = tk.Frame(name_frame, bg='#ffffff')
        name_inner.pack()
        
        name_check = tk.Checkbutton(name_inner, text="生成する", 
                                   variable=self.name_enabled,
                                   bg='#ffffff', fg='#2c3e50',
                                   font=('Arial', 11), selectcolor='#ffffff',
                                   activebackground='#ffffff',
                                   activeforeground='#3498db',
                                   cursor='hand2', padx=5, pady=5)
        name_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(name_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.name_column = tk.StringVar(value="B")
        ttk.Entry(name_inner, textvariable=self.name_column, width=10, 
                 font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # 性別設定（左列）
        self.gender_enabled = tk.BooleanVar(value=False)
        gender_frame = ttk.LabelFrame(left_column, text="性別", 
                                      padding=15, style='Custom.TLabelframe')
        gender_frame.pack(fill="x", pady=8)
        
        gender_inner = tk.Frame(gender_frame, bg='#ffffff')
        gender_inner.pack()
        
        gender_check = tk.Checkbutton(gender_inner, text="生成する", 
                                     variable=self.gender_enabled,
                                     bg='#ffffff', fg='#2c3e50',
                                     font=('Arial', 11), selectcolor='#ffffff',
                                     activebackground='#ffffff',
                                     activeforeground='#3498db',
                                     cursor='hand2', padx=5, pady=5)
        gender_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(gender_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.gender_column = tk.StringVar(value="C")
        ttk.Entry(gender_inner, textvariable=self.gender_column, width=10, 
                 font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # 生年月日設定（左列）
        self.birthdate_enabled = tk.BooleanVar(value=False)
        birthdate_frame = ttk.LabelFrame(left_column, 
                                         text=f"生年月日（高校新入生 - {self.current_year}年度）", 
                                         padding=15, style='Custom.TLabelframe')
        birthdate_frame.pack(fill="x", pady=8)
        
        birthdate_inner = tk.Frame(birthdate_frame, bg='#ffffff')
        birthdate_inner.pack()
        
        birthdate_check = tk.Checkbutton(birthdate_inner, text="生成する", 
                                        variable=self.birthdate_enabled,
                                        bg='#ffffff', fg='#2c3e50',
                                        font=('Arial', 11), selectcolor='#ffffff',
                                        activebackground='#ffffff',
                                        activeforeground='#3498db',
                                        cursor='hand2', padx=5, pady=5)
        birthdate_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(birthdate_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.birthdate_column = tk.StringVar(value="D")
        ttk.Entry(birthdate_inner, textvariable=self.birthdate_column, 
                 width=10, font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # 住所設定（右列）
        self.address_enabled = tk.BooleanVar(value=False)
        address_frame = ttk.LabelFrame(right_column, text="東京都の住所", 
                                       padding=15, style='Custom.TLabelframe')
        address_frame.pack(fill="x", pady=8)
        
        address_inner = tk.Frame(address_frame, bg='#ffffff')
        address_inner.pack()
        
        address_check = tk.Checkbutton(address_inner, text="生成する", 
                                      variable=self.address_enabled,
                                      bg='#ffffff', fg='#2c3e50',
                                      font=('Arial', 11), selectcolor='#ffffff',
                                      activebackground='#ffffff',
                                      activeforeground='#3498db',
                                      cursor='hand2', padx=5, pady=5)
        address_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(address_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.address_column = tk.StringVar(value="E")
        ttk.Entry(address_inner, textvariable=self.address_column, 
                 width=10, font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # 保護者氏名設定（右列）
        self.guardian_name_enabled = tk.BooleanVar(value=False)
        guardian_name_frame = ttk.LabelFrame(right_column, 
                                             text="保護者氏名（学生と同じ苗字）", 
                                             padding=15, style='Custom.TLabelframe')
        guardian_name_frame.pack(fill="x", pady=8)
        
        guardian_name_inner = tk.Frame(guardian_name_frame, bg='#ffffff')
        guardian_name_inner.pack()
        
        guardian_name_check = tk.Checkbutton(guardian_name_inner, text="生成する", 
                                            variable=self.guardian_name_enabled,
                                            bg='#ffffff', fg='#2c3e50',
                                            font=('Arial', 11), selectcolor='#ffffff',
                                            activebackground='#ffffff',
                                            activeforeground='#3498db',
                                            cursor='hand2', padx=5, pady=5)
        guardian_name_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(guardian_name_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.guardian_name_column = tk.StringVar(value="F")
        ttk.Entry(guardian_name_inner, textvariable=self.guardian_name_column, 
                 width=10, font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # 保護者電話番号設定（右列）
        self.guardian_phone_enabled = tk.BooleanVar(value=False)
        guardian_phone_frame = ttk.LabelFrame(right_column, 
                                              text="保護者電話番号", 
                                              padding=15, style='Custom.TLabelframe')
        guardian_phone_frame.pack(fill="x", pady=8)
        
        guardian_phone_inner = tk.Frame(guardian_phone_frame, bg='#ffffff')
        guardian_phone_inner.pack()
        
        guardian_phone_check = tk.Checkbutton(guardian_phone_inner, text="生成する", 
                                             variable=self.guardian_phone_enabled,
                                             bg='#ffffff', fg='#2c3e50',
                                             font=('Arial', 11), selectcolor='#ffffff',
                                             activebackground='#ffffff',
                                             activeforeground='#3498db',
                                             cursor='hand2', padx=5, pady=5)
        guardian_phone_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(guardian_phone_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.guardian_phone_column = tk.StringVar(value="G")
        ttk.Entry(guardian_phone_inner, textvariable=self.guardian_phone_column, 
                 width=10, font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # メールアドレス設定（右列）
        self.email_enabled = tk.BooleanVar(value=False)
        email_frame = ttk.LabelFrame(right_column, 
                                     text="メールアドレス（2から始まる10桁@ドメイン）", 
                                     padding=15, style='Custom.TLabelframe')
        email_frame.pack(fill="x", pady=8)
        
        email_inner = tk.Frame(email_frame, bg='#ffffff')
        email_inner.pack()
        
        email_check = tk.Checkbutton(email_inner, text="生成する", 
                                    variable=self.email_enabled,
                                    bg='#ffffff', fg='#2c3e50',
                                    font=('Arial', 11), selectcolor='#ffffff',
                                    activebackground='#ffffff',
                                    activeforeground='#3498db',
                                    cursor='hand2', padx=5, pady=5)
        email_check.grid(row=0, column=0, sticky="w", padx=5)
        
        tk.Label(email_inner, text="列名:", bg='#ffffff', fg='#2c3e50', 
                font=('Arial', 10)).grid(row=0, column=1, padx=20)
        self.email_column = tk.StringVar(value="H")
        ttk.Entry(email_inner, textvariable=self.email_column, 
                 width=10, font=('Arial', 10)).grid(row=0, column=2, padx=5)
        
        # 実行ボタン（両列の下に配置）
        button_container = tk.Frame(scrollable_frame, bg='#f0f4f8')
        button_container.pack(fill="x", pady=25, padx=15)
        
        generate_btn = tk.Button(button_container, 
                                text="データ生成", 
                                command=self.generate_data,
                                bg='#e74c3c',  # 赤系のアクセントカラー
                                fg='white',
                                font=('Arial', 16, 'bold'),
                                padx=50,
                                pady=15,
                                relief='flat',
                                cursor='hand2',
                                activebackground='#c0392b',
                                activeforeground='white',
                                borderwidth=0)
        generate_btn.pack()
        
        # ホバー効果
        def on_enter(e):
            generate_btn['background'] = '#c0392b'
        def on_leave(e):
            generate_btn['background'] = '#e74c3c'
        
        generate_btn.bind("<Enter>", on_enter)
        generate_btn.bind("<Leave>", on_leave)
        
        # キャンバスとスクロールバーを配置
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def select_file(self):
        filename = filedialog.asksaveasfilename(
            title="Excelファイルを選択または作成",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
    
    def generate_student_id(self):
        """2から始まる7桁の学生番号を生成"""
        return f"2{random.randint(0, 999999):06d}"
    
    def generate_tokyo_address(self):
        """東京都の住所を生成"""
        tokyo_wards = [
            "千代田区", "中央区", "港区", "新宿区", "文京区", "台東区", "墨田区",
            "江東区", "品川区", "目黒区", "大田区", "世田谷区", "渋谷区", "中野区",
            "杉並区", "豊島区", "北区", "荒川区", "板橋区", "練馬区", "足立区", "葛飾区", "江戸川区"
        ]
        
        ward = random.choice(tokyo_wards)
        town = self.fake_ja.city()
        chome = random.randint(1, 9)
        banchi = random.randint(1, 30)
        go = random.randint(1, 20)
        
        return f"東京都{ward}{town}{chome}-{banchi}-{go}"
    
    def generate_birthdate(self):
        """高校新入生の生年月日を生成（同学年）"""
        birth_year = self.current_year - 15
        start_date = datetime(birth_year - 1, 4, 2)
        end_date = datetime(birth_year, 4, 1)
        
        days_between = (end_date - start_date).days
        random_days = random.randint(0, days_between)
        
        birth_date = start_date + timedelta(days=random_days)
        return birth_date.strftime("%Y/%m/%d")
    
    def generate_phone_number(self):
        """日本の電話番号を生成"""
        area_codes = ["03", "090", "080", "070"]
        area_code = random.choice(area_codes)
        return f"{area_code}-{random.randint(1000, 9999)}-{random.randint(1000, 9999)}"
    
    def generate_email(self):
        """2から始まる10桁の数字をユーザー名とするメールアドレスを生成"""
        user_part = f"2{random.randint(0, 999999999):09d}"
        domains = ["example.com", "sample.co.jp", "test.jp", "dummy.com", "school.ac.jp"]
        domain = random.choice(domains)
        return f"{user_part}@{domain}"
    
    def generate_guardian_name(self, student_last_name):
        """保護者氏名を生成（学生と同じ苗字）"""
        guardian_first_names = [
            "太郎", "一郎", "二郎", "三郎", "健一", "誠", "隆", "修", "勇",
            "花子", "美子", "恵子", "洋子", "幸子", "明子", "由美", "真理子", "順子"
        ]
        first_name = random.choice(guardian_first_names)
        return f"{student_last_name} {first_name}"
    
    def generate_data(self):
        # ファイルパスのチェック
        if not self.file_path.get():
            messagebox.showerror("エラー", "Excelファイルを選択してください")
            return
        
        # 少なくとも1つの項目が選択されているか確認
        if not any([
            self.student_id_enabled.get(),
            self.name_enabled.get(),
            self.gender_enabled.get(),
            self.birthdate_enabled.get(),
            self.address_enabled.get(),
            self.guardian_name_enabled.get(),
            self.guardian_phone_enabled.get(),
            self.email_enabled.get()
        ]):
            messagebox.showwarning("警告", "少なくとも1つの項目を選択してください")
            return
        
        try:
            count = min(int(self.data_count.get()), 1000)
            start_row = int(self.start_row.get())
            
            print(f"データ生成開始: {count}件")
            
            # ファイルが存在するか確認
            if os.path.exists(self.file_path.get()):
                wb = openpyxl.load_workbook(self.file_path.get())
            else:
                # 新規ファイルを作成
                wb = Workbook()
                # デフォルトシートを削除
                if 'Sheet' in wb.sheetnames:
                    wb.remove(wb['Sheet'])
            
            # シートを取得または作成
            if self.sheet_name.get() in wb.sheetnames:
                ws = wb[self.sheet_name.get()]
            else:
                ws = wb.create_sheet(self.sheet_name.get())
            
            # データを生成
            for i in range(count):
                row = start_row + i
                student_name = ""
                student_last_name = ""
                
                # 学生番号を生成
                if self.student_id_enabled.get():
                    student_id = self.generate_student_id()
                    cell = ws[f"{self.student_id_column.get().upper()}{row}"]
                    cell.value = student_id
                    print(f"学生番号生成: {student_id}")
                
                # 名前を生成
                if self.name_enabled.get():
                    student_name = self.fake_ja.name()
                    student_last_name = student_name.split()[0]
                    cell = ws[f"{self.name_column.get().upper()}{row}"]
                    cell.value = student_name
                    print(f"名前生成: {student_name}")
                
                # 性別を生成
                if self.gender_enabled.get():
                    gender = random.choice(["男性", "女性"])
                    cell = ws[f"{self.gender_column.get().upper()}{row}"]
                    cell.value = gender
                
                # 生年月日を生成
                if self.birthdate_enabled.get():
                    birthdate = self.generate_birthdate()
                    cell = ws[f"{self.birthdate_column.get().upper()}{row}"]
                    cell.value = birthdate
                
                # 住所を生成
                if self.address_enabled.get():
                    address = self.generate_tokyo_address()
                    cell = ws[f"{self.address_column.get().upper()}{row}"]
                    cell.value = address
                
                # 保護者氏名を生成
                if self.guardian_name_enabled.get():
                    if not student_last_name:
                        temp_name = self.fake_ja.name()
                        student_last_name = temp_name.split()[0]
                    guardian_name = self.generate_guardian_name(student_last_name)
                    cell = ws[f"{self.guardian_name_column.get().upper()}{row}"]
                    cell.value = guardian_name
                
                # 保護者電話番号を生成
                if self.guardian_phone_enabled.get():
                    phone = self.generate_phone_number()
                    cell = ws[f"{self.guardian_phone_column.get().upper()}{row}"]
                    cell.value = phone
                
                # メールアドレスを生成
                if self.email_enabled.get():
                    email = self.generate_email()
                    cell = ws[f"{self.email_column.get().upper()}{row}"]
                    cell.value = email
            
            # ファイルを保存
            wb.save(self.file_path.get())
            wb.close()
            
            print("データ生成完了")
            messagebox.showinfo("完了", f"{count}件のデータ生成が完了しました\n\nファイル: {self.file_path.get()}")
            
        except PermissionError:
            messagebox.showerror("エラー", "ファイルが開かれています。\nExcelファイルを閉じてから再度実行してください。")
        except ValueError as e:
            messagebox.showerror("エラー", f"入力値が不正です: {str(e)}")
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました: {str(e)}\n\n詳細: {type(e).__name__}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    root = tk.Tk()
    app = DummyDataGenerator(root)
    root.mainloop()