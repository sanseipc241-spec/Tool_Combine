import os
import fitz
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32

XL_PAPER_A4 = 9  # Excel constant

# ================= UI =================
class CombineTool:
    def __init__(self, root):
        self.root = root
        root.title("Excel → PDF Combine & Rotate")

        self.files = []

        self.listbox = tk.Listbox(root, width=60, height=15)
        self.listbox.pack(padx=10, pady=5)

        btn_frame = tk.Frame(root)
        btn_frame.pack()

        tk.Button(btn_frame, text="Add Excel Files", command=self.add_files).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Move Up", command=lambda: self.move(-1)).grid(row=0, column=1)
        tk.Button(btn_frame, text="Move Down", command=lambda: self.move(1)).grid(row=0, column=2)
        tk.Button(btn_frame, text="Run", command=self.run).grid(row=0, column=3, padx=5)

    def add_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        for f in files:
            self.files.append(f)
            self.listbox.insert(tk.END, os.path.basename(f))

    def move(self, direction):
        sel = self.listbox.curselection()
        if not sel:
            return
        i = sel[0]
        ni = i + direction
        if 0 <= ni < len(self.files):
            self.files[i], self.files[ni] = self.files[ni], self.files[i]
            txt = self.listbox.get(i)
            self.listbox.delete(i)
            self.listbox.insert(ni, txt)
            self.listbox.select_set(ni)

    # ================= CORE =================
    def run(self):
        if not self.files:
            messagebox.showerror("Error", "No files selected")
            return

        out_dir = os.path.join(os.path.dirname(self.files[0]), "OUTPUT")
        pdf_tmp = os.path.join(out_dir, "TMP_PDF")
        os.makedirs(pdf_tmp, exist_ok=True)

        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        pdf_list = []

        # STEP 1: Excel → A4 PDF (giữ orientation)
        for f in self.files:
            wb = excel.Workbooks.Open(f)
            for ws in wb.Worksheets:
                ps = ws.PageSetup
                ps.PaperSize = XL_PAPER_A4   # ÉP A4
                # KHÔNG set Orientation → giữ ngang/dọc gốc

            pdf_path = os.path.join(
                pdf_tmp, os.path.splitext(os.path.basename(f))[0] + ".pdf"
            )
            wb.ExportAsFixedFormat(0, pdf_path)
            wb.Close(False)
            pdf_list.append(pdf_path)

        excel.Quit()

        # STEP 2: Combine
        combined = fitz.open()
        for p in pdf_list:
            d = fitz.open(p)
            combined.insert_pdf(d)
            d.close()

        combined_path = os.path.join(out_dir, "Combined.pdf")
        combined.save(combined_path)
        combined.close()

        # STEP 3: Rotate (GIỮ NGUYÊN LOGIC CỦA BẠN)
        doc = fitz.open(combined_path)
        for page in doc:
            rect = page.rect
            if rect.width > rect.height:
                page.set_rotation((page.rotation + 270) % 360)

        rotated_path = os.path.join(out_dir, "Rotated_Combined.pdf")
        doc.save(rotated_path)
        doc.close()

        messagebox.showinfo("Done", "Hoan tat combine + rotate")

# ================= MAIN =================
if __name__ == "__main__":
    root = tk.Tk()
    app = CombineTool(root)
    root.mainloop()
