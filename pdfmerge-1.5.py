import os
from tkinter import Tk, Button, Label, Entry, filedialog, Frame, Scrollbar, Canvas, messagebox
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageTk  # 用來顯示縮圖

class PDFMergerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 PDF 合併工具")
        self.root.geometry("800x600")

        self.pdf_files = []

        self.canvas = Canvas(root, borderwidth=0, background="#f8f8f8")
        self.frame = Frame(self.canvas, background="#f8f8f8")
        self.scrollbar = Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((0, 0), window=self.frame, anchor='nw')
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.title_label = Label(self.frame, text="📎 PDF 合併工具 - 請逐一選擇 PDF 並指定頁碼", font=("Arial", 14, "bold"), background="#f8f8f8")
        self.title_label.pack(pady=20)

        self.add_button = Button(self.frame, text="➕ 新增 PDF 檔案", font=("Arial", 12), width=20, command=self.add_pdf)
        self.add_button.pack(pady=10)

        # 設為全部頁面的按鈕
        self.set_all_button = Button(self.frame, text="📚 設為全部頁面", font=("Arial", 12), width=25, command=self.set_all_to_all_pages)
        self.set_all_button.pack(pady=5)

        # 合併按鈕
        self.merge_button = Button(self.frame, text="🔄 合併所有選取頁面", font=("Arial", 12), width=25, command=self.merge_pdfs)
        self.merge_button.pack(pady=20)

    def add_pdf(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file_path in file_paths:
            if file_path:
                file_frame = Frame(self.frame, background="#eef", pady=5)
                file_frame.pack(fill="x", padx=10, pady=5)

                # 使用 grid 佈局
                file_label = Label(file_frame, text=f"📄 {os.path.basename(file_path)}", font=("Arial", 12), background="#eef", anchor="w", width=70)
                file_label.grid(row=0, column=0, columnspan=2, sticky="w")

                page_entry = Entry(file_frame, font=("Arial", 12), width=40)
                page_entry.insert(0, "全部 或 輸入頁碼範圍（如 1-3,5,8-9）")
                page_entry.grid(row=1, column=0, padx=5)

                remove_button = Button(file_frame, text="❌ 移除", font=("Arial", 10), bg="#f66", fg="white",
                                       command=lambda f=file_frame: self.remove_pdf(f))
                remove_button.grid(row=1, column=1, padx=5)

                # 上移、下移按鈕
                up_button = Button(file_frame, text="⬆️ 上移", font=("Arial", 10), command=lambda f=file_frame: self.move_pdf(f, "up"))
                up_button.grid(row=2, column=0, padx=5)

                down_button = Button(file_frame, text="⬇️ 下移", font=("Arial", 10), command=lambda f=file_frame: self.move_pdf(f, "down"))
                down_button.grid(row=2, column=1, padx=5)

                # 頁面縮圖
                preview_label = Label(file_frame, text="預覽", font=("Arial", 10), background="#eef")
                preview_label.grid(row=3, column=0, columnspan=2)

                thumbnail = self.generate_thumbnail(file_path)
                if thumbnail:
                    preview_image = ImageTk.PhotoImage(thumbnail)
                    preview_button = Button(file_frame, image=preview_image, relief="flat", command=lambda path=file_path: self.preview_pdf(path))
                    preview_button.image = preview_image  # 防止圖片被垃圾回收
                    preview_button.grid(row=3, column=2)

                self.pdf_files.append((file_path, page_entry, file_frame))

    def set_all_to_all_pages(self):
        for _, entry, _ in self.pdf_files:
            entry.delete(0, "end")
            entry.insert(0, "全部")

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
        for _, _, frame in self.pdf_files:
            frame.pack_forget()

        # 依照順序重新排版
        for _, _, frame in self.pdf_files:
            frame.pack(fill="x", padx=10, pady=5)

    def generate_thumbnail(self, pdf_path):
        try:
            # 讀取 PDF 檔案
            reader = PdfReader(pdf_path)
            first_page = reader.pages[0]
            image = first_page.to_image()
            thumbnail = image.resize((100, 100))  # 修改為縮圖大小
            return thumbnail
        except Exception as e:
            return None

    def preview_pdf(self, pdf_path):
        # 預覽 PDF 檔案的頁面
        reader = PdfReader(pdf_path)
        first_page = reader.pages[0]
        messagebox.showinfo("預覽頁面", first_page.extract_text())
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
            for pdf_path, entry, _ in self.pdf_files:
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
    root = Tk()
    app = PDFMergerUI(root)
    root.mainloop()
