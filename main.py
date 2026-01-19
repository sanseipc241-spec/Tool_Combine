import os
import shutil
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32

XL_PAPER_A4 = 9  # Excel constant


class CombineTool:
    def __init__(self, root):
        self.root = root
        root.title("Excel -> PDF Combine & Rotate (A4)")

        self.files = []

        self.listbox = tk.Listbox(root, width=70, height=15)
        self.listbox.pack(padx=10, pady=5)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)

        tk.Button(btn_frame, text="Add Excel Files", command=self.add_files).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Move Up", command=lambda: self.move(-1)).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="Move Down", command=lambda: self.move(1)).grid(row=0, column=2, padx=5)
        tk.Button(btn_frame, text="Run", command=self.run).grid(row=0, column=3, padx=5)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Select Excel files",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        for f in files:
            if f not in self.files:
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

            text = self.listbox.get(i)
            self.listbox.delete(i)
            self.listbox.insert(ni, text)
            self.listbox.select_set(ni)

    def run(self):
        if not self.files:
            messagebox.showerror("Error", "No Excel files selected")
            return

        base_dir = os.path.dirname(self.files[0])
        out_dir = os.path.join(base_dir, "OUTPUT")
        pdf_tmp = os.path.join(out_dir, "TMP_PDF")

        os.makedirs(pdf_tmp, exist_ok=True)

        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        pdf_list = []

        try:
            # ===== STEP 1: Excel -> PDF (A4, keep orientation) =====
            for f in self.files:
                wb = excel.Workbooks.Open(f)

                for ws in wb.Worksheets:
                    ps = ws.PageSetup
                    ps.PaperSize = XL_PAPER_A4
                    # KHONG set Orientation -> giu nguyen ngang / doc

                pdf_path = os.path.join(
                    pdf_tmp,
                    os.path.splitext(os.path.basename(f))[0] + ".pdf"
                )

                wb.ExportAsFixedFormat(0, pdf_path)
                wb.Close(False)

                pdf_list.append(pdf_path)

        finally:
            excel.Quit()

        # ===== STEP 2: Combine PDF =====
        combined_path = os.path.join(out_dir, "Combined.pdf")
        combined_doc = fitz.open()

        for p in pdf_list:
            d = fitz.open(p)
            combined_doc.insert_pdf(d)
            d.close()

        combined_doc.save(combined_path)
        combined_doc.close()

        # ===== STEP 3: Rotate landscape pages =====
        rotated_path = os.path.join(out_dir, "Rotated_Combined.pdf")
        doc = fitz.open(combined_path)

        for page in doc:
            rect = page.rect
            if rect.width > rect.height:
                page.set_rotation((page.rotation + 270) % 360)

        doc.save(rotated_path)
        doc.close()

        # ===== STEP 4: CLEAN UP (CHI GIU FILE CUOI) =====
        try:
            if os.path.exists(combined_path):
                os.remove(combined_path)

            if os.path.exists(pdf_tmp):
                shutil.rmtree(pdf_tmp)
        except Exception as e:
            print("Canh bao: khong xoa duoc file tam:", e)

        messagebox.showinfo("Done", "Hoan tat! Chi con Rotated_Combined.pdf")


if __name__ == "__main__":
    root = tk.Tk()
    app = CombineTool(root)
    root.mainloop()
