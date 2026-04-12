import os, ctypes, re, sys
from pathlib import Path
#from tkinter import Tk, Button, Label, Entry, filedialog, Frame, Scrollbar, Canvas, messagebox, Toplevel, StringVar
#from tkinter.ttk import Combobox
from tkinter import Tk, Toplevel, scrolledtext
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox, Canvas
from tkinter import ttk
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageTk  # 用來顯示縮圖
from pdf2image import convert_from_path

#例外處理
import _tkinter
from PyPDF2.errors import PdfReadError

#POPPLER_PATH = r".\poppler-24.08.0\Library\bin"

class PDFMergerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 PDF 合併工具")
        self.root.geometry("1100x750")
        self.root.configure(bg="#f5f5f7") # 現代感的淺灰色背景

        # 設定ttk樣式區塊
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Microsoft JhengHei", 10)) # 統一按鈕樣式
        self.style.configure("Action.TButton", font=("Microsoft JhengHei", 11, "bold")) # 自訂按鈕樣式
        self.style.configure("Flat.TButton", borderwidth=0, focusthickness=0, relief="flat", font=("Microsoft JhengHei", 10))
        self.style.configure("File.TFrame", background="#ffffff") # 自訂訊框樣式
        self.style.configure("My.TRadiobutton", background="#ffffff", font=("Microsoft JhengHei", 10))
        self.flat_btn_style = {
                "fg_color": "#007AFF",
                "hover_color": "#005BB7",
                "corner_radius": 0,
                "border_width": 0,
                "font": ("Microsoft JhengHei", 12, "bold")}
        self.pdf_files = []



        # 1. Header:設定頂部標題區
        self.header = tk.Frame(root, bg="#ffffff", height=80)
        self.header.pack(fill="x", side="top")
        self.label = tk.Label(self.header, text="📄 PDF 合併工具", font=("Microsoft JhengHei", 18, "bold"),
                              bg="#ffffff", fg="#333333")
        self.label.pack(pady=(15,5)) #上方15像素，下方5像素

        

        # 2. Body:主面板區
        self.main_body = tk.Frame(root, bg="#f5f5f7")
        self.main_body.pack(fill="both", expand=True, padx=(10,5), pady=10)

            # 2.1 左側區塊: 可捲動檔案清單
        self.scrollable_frame = ctk.CTkScrollableFrame(self.main_body, corner_radius=15, fg_color="#ffffff",  border_width=1, border_color="#e1e4e8")
        self.scrollable_frame.pack(side="left", expand=True, fill="both")

            # 2.2 右側區塊: 進階設定側邊欄
        self.sidebar_frame = ctk.CTkScrollableFrame(self.main_body, corner_radius=15, fg_color="#ffffff", width=320, border_width=1, border_color="#e1e4e8")
        self.sidebar_frame.pack(side="right", fill="y", padx=(5,0))

        self.sidebar_label = tk.Label(self.sidebar_frame, text="進階設定", font=("Microsoft JhengHei", 13, "bold"), bg="#ffffff", fg="#333333")
        self.sidebar_label.pack()

            # 2.2.1 檔案匯入區塊
        self.import_files_frame = ctk.CTkFrame(self.sidebar_frame, corner_radius=15, fg_color="#ffffff",  border_width=1, border_color="#e1e4e8")
        self.import_files_frame.pack(side="top", fill="x", padx=5, pady=5)
        self.import_files_frame.columnconfigure(0, weight=1)  # 欄位平分
        self.import_files_frame.columnconfigure(1, weight=1)  # 欄位平分

            # 設定頁面標題
        self.import_label = tk.Label(self.import_files_frame, text="批次匯入", font=("Microsoft JhengHei", 11, "bold"), bg="#ffffff", fg="#333333")
        self.import_label.grid(row=0, column=0, columnspan=2, pady=(12,2))


            # 2.2.1.1 匯入路徑區塊
        self.import_path_frame = ctk.CTkFrame(self.import_files_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.import_path_frame.grid(row=1, column=0, columnspan=2, padx=4, pady=(0,2))

        self.inport_path_label = tk.Label(self.import_path_frame, text="1. 選擇資料夾*", font=("Microsoft JhengHei", 10), bg="#ffffff", fg="#333333")
        self.inport_path_label.grid(row=0, column=0, columnspan=2, padx=(10,0), pady=(2,2), sticky="w")
        

            # 輸入路徑欄位
        self.import_path_entry = ttk.Entry(self.import_path_frame, width=30)
        #self.target_folder_entry.grid(row=1, column=2, sticky="w", padx=8, pady=(0,4), ipady=3)
        self.import_path_entry.grid(row=1, column=0, sticky="e", padx=(8,2), pady=(0,0), ipady=2)


            # 選擇路徑按鈕
        self.import_path_button = ttk.Button(self.import_path_frame, text="📂", width=2,
                                     command=lambda: self.set_path_entry(self.import_path_entry))
        #self.set_path_button.grid(row=1, column=1, sticky="e", padx=8, pady=(0,4))
        self.import_path_button.grid(row=1, column=1, padx=(2,8), pady=(0,0), sticky="w")
        

            # 2.2.1.2 關鍵字區塊
        self.keyword_frame = ctk.CTkFrame(self.import_files_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.keyword_frame.grid(row=2, column=0, columnspan=2, padx=4, pady=(0,2))

        self.keyword_label = tk.Label(self.keyword_frame, text="2. 依關鍵字搜尋", font=("Microsoft JhengHei", 10), bg="#ffffff", fg="#333333")
        self.keyword_label.grid(row=0, column=0, columnspan=2, padx=(10,0), pady=(2,2), sticky="w")
        

            # 輸入關鍵字欄位
        self.keyword_entry = ttk.Entry(self.keyword_frame, width=30)
        self.keyword_entry.grid(row=1, column=0, sticky="e", padx=(8,2), pady=(0,0), ipady=2)


            # 關鍵字說明按鈕
        self.hint_button = ttk.Button(self.keyword_frame, text="？", width=2,
                                     command=lambda: messagebox.showinfo("提示", "1.請輸入檔案路徑中包含的關鍵字，並以空格或\",\"區隔。\n2.將匯入選取資料夾中，路徑包含所有關鍵字的檔案。"))
        self.hint_button.grid(row=1, column=1, padx=(2,8), pady=(0,0), sticky="w")

            # 2.2.1.3 分隔線 
        line = tk.Frame(self.import_files_frame, bg="#e1e4e8", height=2)
        line.grid(row=3, column=0, columnspan=2, sticky="ew", padx=20, pady=(8,10))

        
            # 2.2.1.4 匯入/刪除按鈕
        self.import_files_button = ttk.Button(self.import_files_frame, text="匯入", width=12,
                                     command=lambda: self.batch_add_pdfs())
        self.import_files_button.grid(row=4, column=0, padx=(0,6), pady=(2,20), sticky='e')

            #"✅ ❎"
        self.reset_button = ttk.Button(self.import_files_frame, text="重設", width=12,
                                     command=lambda: self.clear_import_entries())
        self.reset_button.grid(row=4, column=1, padx=(6,0), pady=(2,20), sticky='w')

            # 2.2.2 檔案匯出區塊
        self.export_files_frame = ctk.CTkFrame(self.sidebar_frame, corner_radius=15, fg_color="#ffffff",  border_width=1, border_color="#e1e4e8")
        self.export_files_frame.pack(side="top", fill="x", padx=5, pady=5)
        self.export_files_frame.columnconfigure(0, weight=1)
        #self.export_files_frame.columnconfigure(0, weight=1)  # 欄位平分
        #self.export_files_frame.columnconfigure(1, weight=1)  # 欄位平分

            # 設定頁面標題
        self.export_label = tk.Label(self.export_files_frame, text="匯出設定", font=("Microsoft JhengHei", 11, "bold"), bg="#ffffff", fg="#333333")
        self.export_label.grid(row=0, column=0, pady=(12,2), columnspan=2)

            # 2.2.2.1 分隔線
        line = tk.Frame(self.export_files_frame, bg="#e1e4e8", height=1)
        line.grid(row=1, column=0, sticky="ew", padx=20, pady=(8,8), columnspan=2)


            # 2.2.2.2.1 設定合併路徑
        self.merge_path_frame = ctk.CTkFrame(self.export_files_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.merge_path_frame.grid(row=2, column=0, padx=4, pady=(0,0))

        #self.merge_path_frame.columnconfigure((1, 2), weight=1, uniform="group1") # 欄位平分
        self.merge_path_frame.columnconfigure(1, weight=1)
        self.merge_path_frame.columnconfigure(2, weight=3)

        self.merge_path_label = tk.Label(self.merge_path_frame, text="合併至：", font=("Microsoft JhengHei", 10), bg="#ffffff", fg="#333333")
        self.merge_path_label.grid(row=0, column=0, padx=(0,0), sticky="w")

        self.merge_path_option = tk.StringVar(value="A")
        self.merge_r1 = ttk.Radiobutton(self.merge_path_frame, text="預設", variable=self.merge_path_option, value="A", style="My.TRadiobutton")
        self.merge_r1.grid(row=1, column=0, padx=(2,0), pady=(0,0), sticky="w")

        #用一個frame包裝自訂選項
        self.custom_merge_frame = ctk.CTkFrame(self.merge_path_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.custom_merge_frame.grid(row=2, column=0, padx=(2,0), pady=(0,0), sticky="w")

        self.merge_r2 = ttk.Radiobutton(self.custom_merge_frame, text="自訂", variable=self.merge_path_option, value="B", style="My.TRadiobutton")
        self.merge_r2.grid(row=0, column=0, padx=(0,4), pady=(0,0), sticky="w")

        self.custom_merge_path = ttk.Entry(self.custom_merge_frame, width=20)
        self.custom_merge_path.grid(row=0, column=1, sticky="w")
        #self.custom_merge_path.config(state="readonly")

        self.custom_merge_button = ttk.Button(self.custom_merge_frame, text="📂", width=2,
                                     command=lambda: self.set_path_entry(self.custom_merge_path))
        self.custom_merge_button.grid(row=0, column=2, padx=(2,0), sticky="w")
        
        #end_label=tk.Label(self.merge_path_frame, text=" 路徑", font=("Microsoft JhengHei", 10), bg="#ffffff", fg="#333333").grid(row=0, column=3, padx=(0,0))

            # 2.2.2.2.2 設定的編輯/確認的控制區塊
        self.merge_path_control_frame = ctk.CTkFrame(self.export_files_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.merge_path_control_frame.grid(row=2, column=1, padx=(0,10), sticky='n')

        edit_button = ttk.Button(self.merge_path_control_frame, text="✎", width=2, style="Flat.TButton",
                                     command=lambda: self.set_path_entry(self.import_path_entry))
        edit_button.grid(row=0, column=0)

        comfirm_button = ttk.Button(self.merge_path_control_frame, text="✔", width=2, style="Flat.TButton",
                                     command=lambda: self.set_path_entry(self.import_path_entry))
        comfirm_button.grid(row=0, column=1)

            # 2.2.2.3 分隔線
        line = tk.Frame(self.export_files_frame, bg="#e1e4e8", height=1)
        line.grid(row=3, column=0, sticky="ew", padx=20, pady=(8,8), columnspan=2)
                  
            # 2.2.2.4.1 設定分割路徑
        self.split_path_frame = ctk.CTkFrame(self.export_files_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.split_path_frame.grid(row=4, column=0, padx=4, pady=(0,15))

        #self.split_path_frame.columnconfigure((1, 3), weight=1, uniform="group2") # 欄位平分
        #self.split_path_frame.columnconfigure(2, weight=2)

        self.split_path_label = tk.Label(self.split_path_frame, text="分割至：", font=("Microsoft JhengHei", 10), bg="#ffffff", fg="#333333")
        self.split_path_label.grid(row=0, column=0, padx=(0,0), sticky="w")
            
        self.split_path_option = tk.StringVar(value="A")
        self.split_r1 = ttk.Radiobutton(self.split_path_frame, text="預設", variable=self.split_path_option, value="A", style="My.TRadiobutton")
        self.split_r2 = ttk.Radiobutton(self.split_path_frame, text="原檔案路徑", variable=self.split_path_option, value="B", style="My.TRadiobutton")
        self.split_r1.grid(row=1, column=0, padx=(2,0), pady=(0,0), sticky="w")
        self.split_r2.grid(row=2, column=0, padx=(2,0), pady=(0,0), sticky="w")

            #用一個frame包裝自訂選項
        self.custom_split_frame = ctk.CTkFrame(self.split_path_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.custom_split_frame.grid(row=3, column=0, padx=(2,0), pady=(0,0), sticky="w")
        
        self.split_r3 = ttk.Radiobutton(self.custom_split_frame, text="自訂", variable=self.split_path_option, value="C", style="My.TRadiobutton")
        self.split_r3.grid(row=0, column=0, padx=(0,4), pady=(0,0), sticky="w")

        self.custom_split_path = ttk.Entry(self.custom_split_frame, width=20)
        self.custom_split_path.grid(row=0, column=1, sticky="w")
        #self.custom_split_path.config(state="readonly")

        self.custom_split_button = ttk.Button(self.custom_split_frame, text="📂", width=2,
                                     command=lambda: self.set_path_entry(self.custom_split_path))
        self.custom_split_button.grid(row=0, column=2, padx=(2,0), sticky="w")
        #end_label=tk.Label(self.split_path_frame, text=" 路徑", font=("Microsoft JhengHei", 10), bg="#ffffff", fg="#333333").grid(row=0, column=4, padx=(0,0))

            # 2.2.2.4.2 設定的編輯/確認的控制區塊
        self.split_path_control_frame = ctk.CTkFrame(self.export_files_frame, corner_radius=15, fg_color="#ffffff",  border_width=0, border_color="#e1e4e8")
        self.split_path_control_frame.grid(row=4, column=1, padx=(0,10), sticky='n')

        edit_button = ttk.Button(self.split_path_control_frame, text="✎", width=2, style="Flat.TButton",
                                     command=lambda: self.set_path_entry(self.import_path_entry))
        edit_button.grid(row=0, column=0)

        comfirm_button = ttk.Button(self.split_path_control_frame, text="✔", width=2, style="Flat.TButton",
                                     command=lambda: self.set_path_entry(self.import_path_entry))
        comfirm_button.grid(row=0, column=1)

      
            # 2.2.3 設定頁面區塊
        self.set_page_frame = ctk.CTkFrame(self.sidebar_frame, corner_radius=15, fg_color="#ffffff",  border_width=1, border_color="#e1e4e8")
        self.set_page_frame.pack(side="top", fill="x", padx=5, pady=5)
        self.set_page_frame.columnconfigure(0, weight=1)  # 欄位平分
        self.set_page_frame.columnconfigure(1, weight=1)  # 欄位平分

            # 預設的頁面範圍提示
        self.default_range_text = "全部 或 輸入頁碼範圍（如 1-3,5,8-9）"
        
            # 設定頁面標題
        self.page_label = tk.Label(self.set_page_frame, text="快速設定頁碼", font=("Microsoft JhengHei", 11, "bold"), bg="#ffffff", fg="#333333")
        self.page_label.grid(row=0, column=0, columnspan=2, pady=(12,4))

            # 2.2.3.1 
            # 設定下拉選單
        self.combobox = ttk.Combobox(self.set_page_frame, values=["請選擇","全部", "前○頁", "第○頁", "第○-○頁", "第○-倒數第○頁", "後○頁", "最後一頁"], state="readonly", width=11)
        self.combobox.set("請選擇")
        self.combobox.grid(row=1, column=0, padx=8, pady=(0,4), ipady=3, sticky="e")

            # 設定輸入欄位
        vcmd = (self.root.register(self.validate_entry), '%P', '%s') #註冊驗證函式
        self.entry_var = tk.StringVar(value="")
        self.customized_entry = ttk.Entry(
            self.set_page_frame,
            textvariable=self.entry_var,
            width=13,
            validate="key", #在按鍵輸入時觸發驗證
            validatecommand=vcmd #綁定驗證函式
        )
        self.customized_entry.grid(row=1, column=1, padx=8, pady=(0,4) , ipady=3, sticky="w")

            # 綁定下拉選單與輸入欄位
        self.combobox.bind("<<ComboboxSelected>>", self.on_template_change)
        
            # 插入橫線
        #tk.Frame(self.set_page_frame, bg="#e1e4e8", height=3).grid(row=2, column=0, columnspan=2, padx=5, pady=2)


            # 一鍵新增按鈕
        self.add_all_pages_button = ttk.Button(self.set_page_frame, text="新增",
                                     command=lambda: self.add_to_customized_entries(self.entry_var.get())
                                     )
        self.add_all_pages_button.grid(row=3, column=0, padx=8, pady=(4,4), sticky="e")


            # 一鍵設定按鈕
        self.set_all_pages_button = ttk.Button(self.set_page_frame, text="套用",
                                     command=lambda: self.set_customized_entries(self.entry_var.get()) if self.entry_var.get()!="全部" else self.set_all_to_all_pages()
                                     )
        self.set_all_pages_button.grid(row=3, column=1, padx=8, pady=(4,4), sticky="w")


            #一鍵刪除按鈕
        self.clear_all_pages_button = ttk.Button(self.set_page_frame, text="清空",
                                     command=lambda: self.clear_customized_entries()
                                     )
        self.clear_all_pages_button.grid(row=4, column=0, padx=8, pady=(4,10), sticky="e")
    
            #一鍵排序按鈕
        self.sort_all_pages_button = ttk.Button(self.set_page_frame, text="排序",
                                     command=lambda: self.sort_customized_entries()
                                     )
        self.sort_all_pages_button.grid(row=4, column=1, padx=8, pady=(4,10), sticky="w")


            #2.2.4 分割設定區塊
        self.set_split_frame = ctk.CTkFrame(self.sidebar_frame, corner_radius=15, fg_color="#ffffff",  border_width=1, border_color="#e1e4e8")
        self.set_split_frame.pack(side="top", fill="x", padx=5, pady=5)
        self.merge_label = tk.Label(self.set_split_frame, text="分割設定", font=("Microsoft JhengHei", 11, "bold"), bg="#ffffff", fg="#333333")
        self.merge_label.pack(pady=(5,5))

        

            #2.2.5 合併設定區塊
        self.set_merge_frame = ctk.CTkFrame(self.sidebar_frame, corner_radius=15, fg_color="#ffffff",  border_width=1, border_color="#e1e4e8")
        self.set_merge_frame.pack(side="top", fill="x", padx=5, pady=5)
        self.merge_label = tk.Label(self.set_merge_frame, text="合併設定", font=("Microsoft JhengHei", 11, "bold"), bg="#ffffff", fg="#333333")
        self.merge_label.pack(pady=(5,5))
        
        # 3. Footer:設定底部功能區
        self.footer = tk.Frame(root, bg="#ffffff", pady=15, highlightthickness=1, highlightbackground="#dddddd")
        self.footer.pack(fill="x", side="bottom")

            # 3.1 操作列
        self.control_frame = tk.Frame(self.footer, bg="#ffffff")
        self.control_frame.pack()


            # 新增pdf按鈕
        self.add_button = ttk.Button(self.control_frame, text="➕ 新增 PDF",width=12,
                                        command=self.manual_add_pdfs).grid(row=0, column=0, padx=5)

            # 移除全部pdf按鈕
        self.delete_files_button = ttk.Button(self.control_frame, text="➖ 移除所有 PDF", width=15,
                                     command=lambda: self.remove_all_pdfs()).grid(row=0, column=1, padx=5)

            
            # 設為全部頁面的按鈕
        self.set_all_button = ttk.Button(self.control_frame, text="📚 一鍵設為全部", width=15, command=self.set_all_to_all_pages)
        self.set_all_button.grid(row=0, column=2, padx=5)

            # 合併PDF按鈕
        self.merge_button = ttk.Button(self.control_frame, text="🔄 開始合併", width=12,
                                       command=self.merge_pdfs).grid(row=0, column=3, padx=5)

            # 分割PDF按鈕
        self.split_button = ttk.Button(self.control_frame, text="🔄 開始分割", width=12,
                                       command=self.split_pdfs).grid(row=0, column=4, padx=5)

    def manual_add_pdfs(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        self.add_pdfs(file_paths)
    
    def batch_add_pdfs(self):
        str_base_path = self.import_path_entry.get().strip()

        #提升使用者體驗，在沒有設定匯入路徑時，讓使用者手動匯入檔案
        if not str_base_path.strip():
            self.manual_add_pdfs()
            return
        
        #判斷路徑是否存在
        base_path = Path(str_base_path)
        if not base_path.exists():
            messagebox.showerror("匯入失敗", f"路徑{str_base_path}不存在。")
            return
                
        file_paths = []
        keywords = re.split(r'[ ,]+', self.keyword_entry.get())

        #遞迴搜尋所有子資料夾與檔案
        for p in base_path.rglob("*"):
            if p.is_file() and p.suffix.lower() == ".pdf":
                path_str = str(p)

                if all(key in path_str for key in keywords):
                        file_paths.append(path_str)
        if file_paths:
            start_import = self.ask_scrollable_yesno("匯入確認","將批次匯入以下檔案:","\n".join(file_paths))
            if start_import:
                self.add_pdfs(file_paths)
        else:
            messagebox.showwarning("匯入失敗", f"沒有符合條件的檔案")
        

    def add_pdfs(self, file_paths):
        for file_path in file_paths:
            if file_path:
                
                # 2.1.0 讀取pdf檔案資訊
                try:
                    reader = PdfReader(file_path)
                except Exception as e:
                    messagebox.showerror("讀取失敗", f"檔案無法解析：{os.path.basename(file_path)}\n\n可能是檔案毀損或加密格式不支援。\n錯誤原因：{e}")
                    continue
                    
                pages_count = len(reader.pages)
                
                # 2.1.1 建立個別檔案的Frame
                file_frame = tk.Frame(self.scrollable_frame, bg="#ffffff", highlightthickness=1,
                                      highlightbackground="#e1e4e8", pady=10)
                file_frame.pack(fill="x", padx=10, pady=8)
                
                # 2.1.1.1 建立左側選單的Frame(檔名、移動/刪除按鈕)
                menu_frame = tk.Frame(file_frame, bg="#ffffff")
                menu_frame.pack(side="left", fill="both", expand="True", padx=15)
                
                # 建立檔名標籤 
                file_label = tk.Label(menu_frame, text="📄 "+os.path.basename(file_path)+" （共"+str(pages_count)+"頁）",
                         font=("Microsoft JhengHei", 11, "bold"), bg="#ffffff", fg="#0366d6")
                #file_label.pack(anchor="w")
                file_label.grid(row=0, column=0, pady=(0,15), sticky="w")

                #2.1.1.1.1 建立移動、刪除按鈕的Frame
                button_frame = tk.Frame(menu_frame, bg="#ffffff")
                button_frame.grid(row=1, column=0, pady=15, sticky="w")
                
                up_button = ttk.Button(button_frame, text="⬆️ 上移", width=10, command=lambda f=file_frame: self.move_pdf(f, "up"))
                up_button.pack(side="left", padx=5)

                down_button = ttk.Button(button_frame, text="⬇ 下移", width=10, command=lambda f=file_frame: self.move_pdf(f, "down"))
                down_button.pack(side="left", padx=5)

                remove_button = ttk.Button(button_frame, text="❌ 刪除", width=10, command=lambda f=file_frame: self.remove_pdf(f))
                remove_button.pack(side="left", padx=5)

                # 2.1.1.2建立底部輸入範圍的Frame
                entry_frame = tk.Frame(menu_frame, bg="#ffffff")
                entry_frame.grid(row=2, column=0, pady=15, sticky="w")
                #entry_frame.pack(side="bottom", fill="both", expand="True", pady=10)

                page_entry = ttk.Entry(entry_frame, font=("Microsoft JhengHei", 10), width=45)
                page_entry.insert(0, self.default_range_text)
                page_entry.pack(anchor="w")
                

                # 2.1.2 建立右側縮圖的Frame
                thumbnail_frame = tk.Frame(file_frame, bg="#ffffff")
                thumbnail_frame.pack(side="right" ,padx=15)
                
                preview_label = tk.Label(thumbnail_frame, text="預覽", font=("Microsoft JhengHei", 10), bg="#ffffff", fg="#0366d6")
                preview_label.pack(side="top", anchor="n")

                thumbnail = self.generate_thumbnail(file_path)
                if thumbnail:
                    preview_image = ImageTk.PhotoImage(thumbnail)
                    preview_button = ttk.Button(thumbnail_frame, image=preview_image,  command=lambda path=file_path: self.preview_pdf(path))
                    preview_button.image = preview_image  # 防止圖片被垃圾回收
                    preview_button.pack(side="bottom")

                self.pdf_files.append((file_path, page_entry, file_frame, reader))

    def set_path_entry(self, entry):
        folder_path = filedialog.askdirectory(title="請選擇一個資料夾")
        entry.delete(0, "end")
        entry.insert(0, folder_path)
        
    def clear_import_entries(self):
        self.import_path_entry.delete(0, "end")
        self.keyword_entry.delete(0, "end")
        
    def validate_entry(self, new_text, old_text):
        """
            new_text(%P):變動後的字串
            old_text(%s):變動前的字串
        """
        # 取得目前下拉選單選中的樣板類型
        template = self.combobox.get()

        if template == "請選擇":
            return "---"

        elif template == "全部":
            return new_text == "全部"
        
        elif template == "最後一頁":
            return new_text == "最後一頁"

        elif template == "第○頁":
            if new_text.startswith("第") and new_text.endswith("頁"):
                middle_part = new_text[1:-1]
                return middle_part == "" or middle_part.isdigit()

        elif template == "前○頁":
            if new_text.startswith("前") and new_text.endswith("頁"):
                middle_part = new_text[1:-1]
                return middle_part == "" or middle_part.isdigit()

        elif template == "後○頁":
            if new_text.startswith("後") and new_text.endswith("頁"):
                middle_part = new_text[1:-1]
                return middle_part == "" or middle_part.isdigit()
                
        elif template == "第○-○頁":
            if new_text.startswith("第") and new_text.endswith("頁"):
                middle_part = new_text[1:-1]
                return middle_part.count("-") < 2 and all((c.isdigit() or c=="○" or c=="") for c in middle_part.split("-"))

        elif template == "第○-倒數第○頁":
            if new_text.startswith("第") and new_text.endswith("頁"):
                middle_part = new_text[:-1]
                return "-倒數第" in middle_part
        
        return False

    def on_template_change(self, event):
        selected = self.combobox.get()

        if selected == "請選擇":
            self.customized_entry.config(state="readonly")
            self.entry_var.set("")
        elif selected=="最後一頁":
            self.customized_entry.config(state="readonly")
            self.entry_var.set("最後一頁")
        else:
            self.customized_entry.config(state="normal")
            self.entry_var.set(selected)


    def set_customized_entries(self, range_text):
        if range_text=="":
            return
        
        warning = False
        value_error = False
        other_error = False
        message1 = "輸入含有不合法字元：\n可能是頁數未設定或頁數為空值。\n\n錯誤原因：\n"
        message2 = "出現非預期錯誤\n\n錯誤原因:\n"
        for _, entry, _,reader in self.pdf_files:
            try:
                #將輸入範圍轉換成指定格式
                pages_count = len(reader.pages)
                new_text = self.match_range(range_text, pages_count)
                entry.delete(0, "end")
                if new_text != "":
                    entry.insert(0, new_text)
                    
            except ValueError as e:
                value_error = True
                message1 += str(e)
            except Exception as e:
                other_error = True
                message2 += str(e)
                
        message = ""
        if value_error:
            message+=message1
        if other_error:
            message+=message2

        if value_error or other_error:
            messagebox.showerror("提示", message)

  
    def add_to_customized_entries(self, range_text):
        if range_text=="":
            return
        
        warning = False
        value_error = False
        other_error = False
        message1 = "輸入含有不合法字元：\n可能是頁數未設定或頁數為空值。\n\n錯誤原因：\n"
        message2 = "出現非預期錯誤\n\n錯誤原因:\n"
        for _, entry, _,reader in self.pdf_files:
            try:
                pages_count = len(reader.pages)

                #將原輸入範圍轉換成指定格式
                original_text = entry.get()
                original_ranges = [self.match_range(item.strip(), pages_count) for item in original_text.split(",")]
                original_ranges = [item for item in original_ranges if item.strip()]
                matched_original_text = ",".join(original_ranges)

                #將新輸入範圍轉換成指定格式
                matched_new_text = self.match_range(range_text, pages_count)

                entry.delete(0, "end")
                
                #若轉換格式後的原輸入範圍不為空，且原輸入範圍不為提示
                if matched_original_text and original_text!=self.default_range_text:
                    entry.insert(0, matched_original_text+","+matched_new_text)

                #若轉換格式後的原輸入範圍為空，且新輸入範圍不為空
                elif matched_new_text != "":
                    entry.insert(0, matched_new_text)
                    
            except ValueError as e:
                value_error = True
                message1 += str(e)
            except Exception as e:
                other_error = True
                message2 += str(e)
                
        message = ""
        if value_error:
            message+=message1
        if other_error:
            message+=message2

        if value_error or other_error:
            messagebox.showerror("提示", message)

    def clear_customized_entries(self):
        for _, entry, _,_ in self.pdf_files:
            entry.delete(0, "end")
        return

    def sort_customized_entries(self):
        pass


    def set_all_to_all_pages(self):
        for _, entry, _,_ in self.pdf_files:
            entry.delete(0, "end")
            entry.insert(0, "全部")

    def match_range(self, range_text, pages_count):
        try:
            range_text = range_text.strip()
            
            #將輸入範圍轉換成指定格式
            pattern = r'(全部|[0-9-]+)'
            matches = re.findall(pattern, range_text)#設定匹配的正則表達式
            pages_locator = ''.join(matches)

            if re.fullmatch(r"前\d+頁", range_text):

                #計算結束頁，並使其不超過總頁數
                end_page = min(int(pages_locator), pages_count)

                if end_page > 1:
                    return "1-"+str(end_page)
                elif end_page == 1:
                    return str(end_page)
                else:
                    return ""

            elif re.fullmatch(r"後\d+頁", range_text):

                #計算起始頁，並使其不低於1頁
                start_page = max(1, pages_count-int(pages_locator)+1)

                if start_page < pages_count:
                    return str(start_page)+"-"+str(pages_count)
                elif start_page == pages_count:
                    return str(pages_count)
                else:
                    return ""

            elif re.fullmatch(r"第\d+(-\d+)?頁", range_text):
                pages = [int(item) for item in pages_locator.split("-")]
                
                #設定多頁
                if len(pages)==2:
                    sorted_pages = sorted(pages)
                    
                    if max(pages)<1 or min(pages)>pages_count:
                        return ""
                    else:
                        mod_pages = [str( min(pages_count, max(1, item)) ) for item in pages]
                        return "-".join(mod_pages) if mod_pages[0]!=mod_pages[1] else mod_pages[0]
                #設定單頁
                elif len(pages)==1:
                    return str(pages[0]) if (pages[0]>0 and pages[0]<=pages_count) else ""

                #非預期設定
                else:
                    return ""
                
            elif re.fullmatch(r"第\d+-倒數第\d+頁", range_text):
                pages = re.findall(r"\d+", range_text)
                end_page = str(pages_count - int(pages[1]) + 1) if pages_count>int(pages[1]) else "1"
                return pages[0]+"-"+end_page

            elif range_text.strip() == "最後一頁":
                return str(pages_count)
            
            elif range_text.strip() == "全部":
                return "全部"

            # 若為合法且轉換後的輸入，直接return原輸入(已去除多餘空白)
            elif re.fullmatch(r"\d+-\d+",range_text) or re.fullmatch(r"\d+",range_text):
                return range_text

            else:
                return ""

        #使用者輸入非法範圍(如○,空值)時報錯。為使設定多個檔案時只報一次錯誤，raise exception到函式外層處理。
        except ValueError as e:
            raise ValueError(str(e)+"\n")
            print(e)
        except Exception as e:
            raise Exception(str(e)+"\n")
            print(e)
            
    def move_pdf(self, file_frame, direction):
        # 找出被點擊的檔案在 list 中的索引
        idx = next((i for i, item in enumerate(self.pdf_files) if item[2] == file_frame), None)
        if idx is not None:
            if direction == "up" and idx > 0:
                self.pdf_files[idx], self.pdf_files[idx - 1] = self.pdf_files[idx - 1], self.pdf_files[idx]
            elif direction == "down" and idx < len(self.pdf_files) - 1:
                self.pdf_files[idx], self.pdf_files[idx + 1] = self.pdf_files[idx + 1], self.pdf_files[idx]

            # 更新 GUI 以反映改動
            self.update_ui()

    def update_ui(self):
        # 清除舊 UI
        for _, _, frame, _ in self.pdf_files:
            frame.pack_forget()

        # 依照順序重新排版
        for _, _, frame, _ in self.pdf_files:
            frame.pack(fill="x", padx=10, pady=5)

    def get_poppler_path(self):
        base = getattr(sys, '_MEIPASS', None)

        if base:  # exe 環境
            return os.path.join(base, "poppler", "Library", "bin")
        else:  # 開發環境
            return os.path.join(os.path.dirname(__file__), "poppler-24.08.0", "Library", "bin")


    def generate_thumbnail(self, pdf_path):
        try:
            POPPLER_PATH = self.get_poppler_path()
            images = convert_from_path(pdf_path, first_page=1, last_page=1, size=(100, 100), poppler_path=POPPLER_PATH)
            if images:
                return images[0]
        except Exception as e:
            print("縮圖錯誤:", e)
        return None

    def preview_pdf(self, pdf_path):
        try:
            # 用高 DPI 轉圖片，但不指定尺寸，保留原始大小比例
            POPPLER_PATH = self.get_poppler_path()
            images = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=200, poppler_path=POPPLER_PATH)
            if not images:
                messagebox.showwarning("預覽失敗", "無法產生預覽圖片。")
                return

            top = Toplevel(self.root)
            top.title("📄 預覽放大圖")

            # 建立 Canvas 搭配 Scrollbar 讓圖片可滾動
            canvas = Canvas(top)
            canvas.pack(side="left", fill="both", expand=True)

            scrollbar = ttk.Scrollbar(top, command=canvas.yview)
            scrollbar.pack(side="right", fill="y")
            canvas.configure(yscrollcommand=scrollbar.set)

            # 圖片轉換成 Tk image 並顯示
            enlarged_image = images[0]
            tk_image = ImageTk.PhotoImage(enlarged_image)

            canvas.image = tk_image  # 防止回收
            canvas.create_image(0, 0, image=tk_image, anchor="nw")

            # 設定 scroll 區域範圍
            canvas.config(scrollregion=canvas.bbox("all"))

            # 可選：限制初始視窗大小（若圖片很大）
            top.geometry("800x1000")

        except Exception as e:
            messagebox.showerror("預覽錯誤", f"無法顯示預覽：{e}")
        
    def parse_pages(self, input_str, max_page):
        input_str = input_str.strip()
        pages = []
        
        for part in input_str.split(","):
            part = part.strip()
            if "-" in part:
                try:
                    start, end = map(int, part.split("-"))
                    pages.extend(range(start-1, end)) if start<end else pages.extend(range(start-1, end-2, -1))
                except ValueError:
                    continue
            elif "全部" in part:
                try:
                    pages.extend(range(max_page))
                except Exception:
                    continue
            else:
                try:
                    pages.append(int(part)-1)
                except ValueError:
                    continue
        return pages
            
        

    def remove_pdf(self, file_frame):
        self.pdf_files = [item for item in self.pdf_files if item[2] != file_frame]
        file_frame.destroy()

    def remove_all_pdfs(self):
        result = False
        if self.pdf_files:
            result = messagebox.askyesno("移除確認", "是否移除所有pdf？")
    
        if result:
            for _,_,file_frame,_ in self.pdf_files:
                file_frame.destroy()
            self.pdf_files = []

    def merge_pdfs(self):
        writer = PdfWriter()
        try:
            #無檔案時直接離開
            if not self.pdf_files:
                return

            for pdf_path, entry, _, reader in self.pdf_files:
                pages_count = len(reader.pages)
                page_input = entry.get()
                selected_pages = self.parse_pages(page_input, pages_count)
                for i in selected_pages:
                    if 0 <= i < pages_count:
                        writer.add_page(reader.pages[i])
        except Exception as e:
            messagebox.showerror("錯誤", f"合併時發生錯誤：{e}")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_path:
            with open(output_path, "wb") as out_file:
                writer.write(out_file)
            messagebox.showinfo("完成", f"✅ 已成功合併 PDF！\n儲存於：\n{output_path}")

    def split_pdfs(self):
        def get_unique_output_path(original_path, suffix="_split"):
            """
            original_path: 原檔案路徑 (str 或 Path)
            suffix: 想要加在原檔名後的字串
            """
            p = Path(original_path)
            # 組合初始的輸出路徑：原資料夾 / (原檔名 + 字尾 + .pdf)
            folder = p.parent
            base_name = p.stem  # 取得不含副檔名的檔名
            ext = p.suffix     # 取得副檔名 (.pdf)
    
            # 初始目標路徑
            target_path = folder / f"{base_name}{suffix}{ext}"
    
            counter = 1
            # 如果檔案已經存在，就進入迴圈尋找下一個可用的序號
            while target_path.exists():
                target_path = folder / f"{base_name}{suffix}({counter}){ext}"
                counter += 1
        
            return target_path
        
        try:
            #無檔案時直接離開
            if not self.pdf_files:
                return
            
            writers = []
            source_paths = []
            for pdf_path, entry, _, reader in self.pdf_files:
                need_split = False

                writer = PdfWriter()
                pages_count = len(reader.pages)
                pages_input= entry.get()
                selected_pages = self.parse_pages(pages_input, pages_count)
                for i in selected_pages:
                    if 0 <= i < pages_count:
                        writer.add_page(reader.pages[i])
                        need_split = True

                # 只匯出要分割的檔案
                if need_split:
                    writers.append(writer)
                    source_paths.append(Path(pdf_path))

            # 依匯出設定輸出
            # 由使用者選擇匯出路徑
            output_paths = []
            if self.split_path_option.get()=="A":
                output_path = filedialog.askdirectory(title="請選擇匯出路徑")
                if output_path:
                    output_paths.extend( [ Path(output_path) ] * len(writers) )
                else:
                    return
            # 匯出至原檔案資料夾    
            elif self.split_path_option.get()=="B":
                output_paths = [ path.parent for path in source_paths] # 取得來原檔案所屬資料夾

            # 匯出至使用者自訂路徑
            else:
                if Path(self.custom_split_path.get()).exists():
                    output_paths.extend( [self.custom_split_path.get()] * len(writers) )
                else:
                    messagebox.showerror("匯出錯誤", "自訂路徑不存在。\n")
                    return
        except Exception as e:
            messagebox.showerror("錯誤", f"分割時發生錯誤：{e}")
            return

        #再次向使用者確認是否分割
        result = messagebox.askyesno("分割提醒", "是否開始分割?")
        if not result:
            return

            # 依序匯出檔案
        messages = []
        for writer, output_path, source_path in zip(writers, output_paths, source_paths):
            try:
                if output_path:
                    #決定輸出的檔名
                    file_path = get_unique_output_path(str(output_path)+"/"+str(source_path.name))

                    with open(file_path, "wb") as out_file:
                        writer.write(out_file)
                        messages.append(f"【成功】分割{source_path}，儲存於：{output_path}")
            except Exception as e:
                messages.append(f"【失敗】分割{source_path}")

        self.show_scrollable_info("完成PDF分割", "分割結果:", "\n".join(messages))
        return
                    
        

    def ask_scrollable_yesno(self, title, header, message):
        """
        建立一個帶有捲軸的『是/否』選擇視窗
        回傳值: True (使用者點擊『是』), False (點擊『否』或直接關閉)
        """
        # 建立一個變數來儲存使用者的選擇
        self._choice_result = False 

        # 建立子視窗
        msg_win = tk.Toplevel(self.root)
        msg_win.title(title)
        msg_win.geometry("700x450")
    
        # 設定為模態視窗 (強制聚焦)
        msg_win.grab_set()

        # 1. 標題與說明
        tk.Label(msg_win, text=header, font=("Microsoft JhengHei", 11, "bold"), 
             pady=10, fg="#333333").pack()

        # 2. 帶捲軸的內容區域
        text_area = scrolledtext.ScrolledText(msg_win, wrap=tk.WORD, width=60, height=15)
        text_area.insert(tk.INSERT, message)
        text_area.config(state='disabled') # 設為唯讀
        text_area.pack(padx=20, pady=10, fill="both", expand=True)

        # 3. 按鈕區 (水平排列)
        btn_frame = tk.Frame(msg_win, pady=15)
        btn_frame.pack(fill="x")

        def on_yes():
            self._choice_result = True
            msg_win.destroy()

        def on_no():
            self._choice_result = False
            msg_win.destroy()

        # 是 (Yes) 按鈕
        yes_btn = ttk.Button(btn_frame, text="是 (Yes)", width=12, command=on_yes)
        yes_btn.pack(side="left", expand=True, padx=(50, 5))

        # 否 (No) 按鈕
        no_btn = ttk.Button(btn_frame, text="否 (No)", width=12, command=on_no)
        no_btn.pack(side="left", expand=True, padx=(5, 50))

        # --- 關鍵步驟：等待視窗關閉 ---
        self.root.wait_window(msg_win)
    
        return self._choice_result
    
    def show_scrollable_info(self, title, header, message):
        # 建立子視窗
        msg_win = tk.Toplevel(self.root)
        msg_win.title(title)
        msg_win.geometry("700x450")
    
        # 設定為模態視窗 (強制聚焦)
        msg_win.grab_set()

        # 1. 標題與說明
        tk.Label(msg_win, text=header, font=("Microsoft JhengHei", 11, "bold"), 
             pady=10, fg="#333333").pack()

        # 2. 帶捲軸的內容區域
        text_area = scrolledtext.ScrolledText(msg_win, wrap=tk.WORD, width=60, height=15)
        text_area.insert(tk.INSERT, message)
        text_area.config(state='disabled') # 設為唯讀
        text_area.pack(padx=20, pady=10, fill="both", expand=True)

        return
    

if __name__ == "__main__":

    # 解決低解析度問題
    # 1. 找到 Windows 負責控制「縮放意識 (DPI Awareness)」的 DLL 檔案
    # 2. 呼叫 SetProcessDpiAwareness 這個 C 語言函式
    # 參數 1 代表「支援系統縮放，但不模糊」
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = Tk()
    app = PDFMergerUI(root)
    root.mainloop()
