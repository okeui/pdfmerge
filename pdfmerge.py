import os
from tkinter import Tk, Button, Label, Entry, filedialog, Frame, Scrollbar, Canvas, messagebox
from PyPDF2 import PdfReader, PdfWriter

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

        self.merge_button = Button(self.frame, text="🔄 合併所有選取頁面", font=("Arial", 12), width=25, command=self.merge_pdfs)
        self.merge_button.pack(pady=20)

    def add_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            # 顯示檔名
            file_label = Label(self.frame, text=f"📄 {os.path.basename(file_path)}", anchor="w", font=("Arial", 12), background="#f8f8f8")
            file_label.pack(fill="x", padx=20, pady=2)

            # 建立頁碼輸入欄
            page_entry = Entry(self.frame, font=("Arial", 12), width=50)
            page_entry.insert(0, "全部 或 輸入頁碼範圍（如 1-3,5,8-9）")
            page_entry.pack(padx=20, pady=5)

            self.pdf_files.append((file_path, page_entry))

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

    def merge_pdfs(self):
        writer = PdfWriter()

        try:
            for pdf_path, entry in self.pdf_files:
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
