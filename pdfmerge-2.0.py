import os
import ctypes
import re
#from tkinter import Tk, Button, Label, Entry, filedialog, Frame, Scrollbar, Canvas, messagebox, Toplevel, StringVar
#from tkinter.ttk import Combobox
from tkinter import Tk, Toplevel
import tkinter as tk
from tkinter import filedialog, messagebox, Canvas
from tkinter import ttk
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageTk  # 用來顯示縮圖
from pdf2image import convert_from_path

POPPLER_PATH = r".\poppler-24.08.0\Library\bin" 

class PDFMergerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 PDF 合併工具")
        self.root.geometry("950x600")
        self.root.configure(bg="#f5f5f7") # 現代感的淺灰色背景

        # 設定ttk樣式區塊
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Microsoft JhengHei", 10)) # 統一按鈕樣式
        self.style.configure("Action.TButton", font=("Microsoft JhengHei", 11, "bold")) # 自訂按鈕樣式
        self.style.configure("File.TFrame", background="#ffffff") # 自訂訊框樣式

        self.pdf_files = []



        # 1. Header:設定頂部標題區
        self.header = tk.Frame(root, bg="#ffffff", height=80)
        self.header.pack(fill="x", side="top")
        self.label = tk.Label(self.header, text="📄 PDF 合併工具", font=("Microsoft JhengHei", 18, "bold"),
                              bg="#ffffff", fg="#333333")
        self.label.pack(pady=(15,5)) #上方15像素，下方5像素

        

        # 2. Body:主面板區，用Canvas模擬可捲動的容器
        self.container = tk.Frame(root, bg="#f5f5f7")
        self.container.pack(fill="both", expand=True, padx=20, pady=10)

            # 設定元件邊框及焦點外框大小為0
        self.canvas = Canvas(self.container, borderwidth=0, background="#f5f5f7", highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        
        self.scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
        self.scrollbar.configure(command=self.canvas.yview) # 拉動scrollbar時更新canvas的捲動位置
        self.canvas.configure(yscrollcommand=self.scrollbar.set) # canvas垂直捲動時通知scrollbar更新位置
        self.scrollbar.pack(side="right", fill="y")

            # 將frame加入canvas，貼齊canvas左上角
        self.scrollable_frame = tk.Frame(self.canvas, background="#f5f5f7")
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor='nw')
            # 監聽frame的<Configure>事件，位置或大小改變時更新canvas的可捲動範圍
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        

        
        # 3. Footer:設定底部功能區
        self.footer = tk.Frame(root, bg="#ffffff", pady=15, highlightthickness=1, highlightbackground="#dddddd")
        self.footer.pack(fill="x", side="bottom")

            # 3.1 操作列
        self.control_frame = tk.Frame(self.footer, bg="#ffffff")
        self.control_frame.pack()

            # 下拉選單與一鍵設定
        self.combobox = ttk.Combobox(self.control_frame, values=["第○頁", "第○-○頁", "前○頁", "後○頁", "最後一頁"], state="readonly", width=12)
        self.combobox.set("第○頁")
        self.combobox.grid(row=0, column=0, padx=5, ipady=3)

        vcmd = (self.root.register(self.validate_entry), '%P', '%s') #註冊驗證函式
        self.entry_var = tk.StringVar(value="第○頁")
        self.customize_entry = ttk.Entry(
            self.control_frame,
            textvariable=self.entry_var,
            width=15,
            validate="key", #在按鍵輸入時觸發驗證
            validatecommand=vcmd #綁定驗證函式
        )
        self.customize_entry.grid(row=0, column=1, padx=5, ipady=3)

        # 綁定下拉選單與輸入欄位
        self.combobox.bind("<<ComboboxSelected>>",lambda event: self.entry_var.set(self.combobox.get()))
        
        self.set_button = ttk.Button(self.control_frame, text="📚 一鍵設定",
                                     command=lambda: self.set_customized_page(self.entry_var.get())
                                     )
        self.set_button.grid(row=0, column=2, padx=5)

            # 設為全部頁面的按鈕
        self.set_all_button = ttk.Button(self.control_frame, text="📚 一鍵設為全部", width=15, command=self.set_all_to_all_pages)
        self.set_all_button.grid(row=0, column=3, padx=5)


            # 新增、合併PDF按鈕
        self.add_button = ttk.Button(self.control_frame, text="➕ 新增 PDF",width=12,
                                     command=self.add_pdf).grid(row=0, column=4, padx=5)
        self.merge_button = ttk.Button(self.control_frame, text="🔄 開始合併", width=12,
                                       command=self.merge_pdfs).grid(row=0, column=5, padx=5)

    def validate_entry(self, new_text, old_text):
        """
            new_text(%P):變動後的字串
            old_text(%s):變動前的字串
        """
        # 取得目前下拉選單選中的樣板類型
        template = self.combobox.get()

        if template == "最後一頁":
            return new_text == "最後一頁"

        if template == "第○頁":
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
        
        return False

    def add_pdf(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file_path in file_paths:
            if file_path:
                # 2.0 讀取pdf檔案資訊
                reader = PdfReader(file_path)
                pages_count = len(reader.pages)
                
                # 2.1 建立個別檔案的Frame
                file_frame = tk.Frame(self.scrollable_frame, bg="#ffffff", highlightthickness=1,
                                      highlightbackground="#e1e4e8", pady=10)
                file_frame.pack(fill="x", padx=10, pady=8)

                # 2.1.1 建立左側選單的Frame(檔名、移動/刪除按鈕)
                menu_frame = tk.Frame(file_frame, bg="#ffffff")
                menu_frame.pack(side="left", fill="both", expand="True", padx=15)
                
                # 建立檔名標籤 
                file_label = tk.Label(menu_frame, text="📄 "+os.path.basename(file_path)+" （共"+str(pages_count)+"頁）",
                         font=("Microsoft JhengHei", 11, "bold"), bg="#ffffff", fg="#0366d6")
                #file_label.pack(anchor="w")
                file_label.grid(row=0, column=0, pady=(0,15), sticky="w")

                #2.1.1.1 建立移動、刪除按鈕的Frame
                button_frame = tk.Frame(menu_frame, bg="#ffffff")
                button_frame.grid(row=1, column=0, pady=15, sticky="w")
                #button_frame.pack(anchor="w", pady=3)
                
                up_button = ttk.Button(button_frame, text="⬆️ 上移", width=10, command=lambda f=file_frame: self.move_pdf(f, "up"))
                up_button.pack(side="left", padx=5)

                down_button = ttk.Button(button_frame, text="⬇ 下移", width=10, command=lambda f=file_frame: self.move_pdf(f, "down"))
                down_button.pack(side="left", padx=5)

                remove_button = ttk.Button(button_frame, text="❌ 刪除", width=10, command=lambda f=file_frame: self.remove_pdf(f))
                remove_button.pack(side="left", padx=5)

                # 2.1.2建立底部輸入範圍的Frame
                entry_frame = tk.Frame(menu_frame, bg="#ffffff")
                entry_frame.grid(row=2, column=0, pady=15, sticky="w")
                #entry_frame.pack(side="bottom", fill="both", expand="True", pady=10)
                
                page_entry = ttk.Entry(entry_frame, font=("Microsoft JhengHei", 10), width=45)
                page_entry.insert(0, "全部 或 輸入頁碼範圍（如 1-3,5,8-9）")
                page_entry.pack(anchor="w")
                

                # 2.2 建立右側縮圖的Frame
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

    def set_customized_page(self, range_text):
        for _, entry, _,reader in self.pdf_files:
            #將輸入範圍轉換成指定格式
            pages_count = len(reader.pages)
            new_text = self.match_range(range_text, pages_count)
            entry.delete(0, "end")
            entry.insert(0, new_text)

    def set_all_to_all_pages(self):
        for _, entry, _,_ in self.pdf_files:
            entry.delete(0, "end")
            entry.insert(0, "全部")
            
    def match_range(self, range_text, pages_count):
        #將輸入範圍轉換成指定格式
        pattern = r'(全部|[0-9-]+)'
        matches = re.findall(pattern, range_text)#設定匹配的正則表達式
        pages_locator = ''.join(matches)

        if range_text.startswith("前") and range_text.endswith("頁"):
            return "1" if pages_count==1 else "1-"+str(min(int(pages_locator), pages_count))
        elif range_text.startswith("後") and range_text.endswith("頁"):
            return "1-"+str(pages_count-int(pages_locator)+1) if int(pages_locator) < (pages_count-1) else "1"
        elif range_text.startswith("第") and range_text.endswith("頁"):
            if "-" in range_text:
                pages = [max(1,str(item)) for item in page_locator.split("-")].sort()
                return "-".join(pages) if pages[0]!=pages[1] else str(pages[0])
            else:
                return str(max(1,min(int(pages_locator), page_count)))
        elif range_text == "最後一頁":
            return str(pages_count)
        elif range_text == "全部":
            return "全部"
        else:
            return "全部"

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


    def generate_thumbnail(self, pdf_path):
        try:
            images = convert_from_path(pdf_path, first_page=1, last_page=1, size=(100, 100), poppler_path=POPPLER_PATH)
            if images:
                return images[0]
        except Exception as e:
            print("縮圖錯誤:", e)
        return None

    def preview_pdf(self, pdf_path):
        try:
            # 用高 DPI 轉圖片，但不指定尺寸，保留原始大小比例
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
        if not input_str.strip() or input_str.lower().startswith("全部"):
            return list(range(max_page))  # 全部頁面（0-indexed）

        pages = set()
        for part in input_str.split(","):
            part = part.strip()
            if "-" in part:
                try:
                    start, end = map(int, part.split("-"))
                    pages.update(range(start - 1, end))
                except ValueError:
                    continue
            else:
                try:
                    pages.add(int(part) - 1)
                except ValueError:
                    continue
        return sorted(pages)

    def remove_pdf(self, file_frame):
        self.pdf_files = [item for item in self.pdf_files if item[2] != file_frame]
        file_frame.destroy()

    def merge_pdfs(self):
        writer = PdfWriter()
        try:
            for pdf_path, entry, _, _ in self.pdf_files:
                reader = PdfReader(pdf_path)
                page_input = entry.get()
                selected_pages = self.parse_pages(page_input, len(reader.pages))
                for i in selected_pages:
                    if 0 <= i < len(reader.pages):
                        writer.add_page(reader.pages[i])
        except Exception as e:
            messagebox.showerror("錯誤", f"合併時發生錯誤：{e}")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_path:
            with open(output_path, "wb") as out_file:
                writer.write(out_file)
            messagebox.showinfo("完成", f"✅ 已成功合併 PDF！\n儲存於：\n{output_path}")

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
