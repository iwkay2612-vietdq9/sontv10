import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import xlwings as xw

class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lọc Excel & Giữ Định Dạng (Python Version)")
        self.root.geometry("700x600")
        
        # Variables
        self.source_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.filter_rows = []
        self.mode = tk.IntVar(value=1) # 1=Merge, 2=Split
        
        # === UI LAYOUT ===
        
        # File Source
        frame_file = tk.LabelFrame(root, text="File Nguồn", padx=10, pady=10)
        frame_file.pack(fill="x", padx=10, pady=5)
        
        tk.Entry(frame_file, textvariable=self.source_file, width=50).pack(side="left", padx=5)
        tk.Button(frame_file, text="Chọn File", command=self.browse_file).pack(side="left")

        # Output Folder (Optional, default to source folder)
        # We'll just save to same folder or ask. Let's save to 'Output' subfolder.
        
        # Filters
        frame_filter = tk.LabelFrame(root, text="Bộ Lọc (AND giữa cột, OR trong giá trị)", padx=10, pady=10)
        frame_filter.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Header for filters
        header_frame = tk.Frame(frame_filter)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="Tên Cột (Header)", width=25, anchor="w").pack(side="left")
        tk.Label(header_frame, text="Giá trị (cách nhau dấu phẩy)", anchor="w").pack(side="left", padx=10)
        
        self.filter_container = tk.Frame(frame_filter)
        self.filter_container.pack(fill="both", expand=True)
        
        # Add 5 initial rows
        for _ in range(5):
            self.add_filter_row()
            
        tk.Button(frame_filter, text="+ Thêm dòng lọc", command=self.add_filter_row).pack(anchor="e", pady=5)
        
        # Options
        frame_opts = tk.LabelFrame(root, text="Tùy Chọn", padx=10, pady=10)
        frame_opts.pack(fill="x", padx=10, pady=5)
        
        tk.Radiobutton(frame_opts, text="1) Gộp tất cả vào 1 file", variable=self.mode, value=1, command=self.toggle_split_settings).pack(anchor="w")
        tk.Radiobutton(frame_opts, text="2) Tách TỰ ĐỘNG theo cột", variable=self.mode, value=2, command=self.toggle_split_settings).pack(anchor="w")
        
        self.frame_split = tk.Frame(frame_opts)
        self.frame_split.pack(fill="x", pady=5)
        tk.Label(self.frame_split, text="Cột để tách file:").pack(side="left")
        self.entry_split_col = tk.Entry(self.frame_split, width=30)
        self.entry_split_col.pack(side="left", padx=5)
        
        self.toggle_split_settings()

        # Run Button
        tk.Button(root, text="CHẠY XỬ LÝ", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), command=self.run_process).pack(pady=20, ipadx=20, ipady=5)
        
        self.status_label = tk.Label(root, text="Sẵn sàng", fg="blue")
        self.status_label.pack()

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.source_file.set(filename)

    def add_filter_row(self):
        row_frame = tk.Frame(self.filter_container)
        row_frame.pack(fill="x", pady=2)
        
        ent_col = tk.Entry(row_frame, width=25)
        ent_col.pack(side="left")
        
        ent_val = tk.Entry(row_frame, width=45)
        ent_val.pack(side="left", padx=10)
        
        self.filter_rows.append((ent_col, ent_val))

    def toggle_split_settings(self):
        if self.mode.get() == 2:
            self.entry_split_col.config(state="normal")
        else:
            self.entry_split_col.config(state="disabled")

    def log(self, msg):
        self.status_label.config(text=msg)
        self.root.update()

    def run_process(self):
        source = self.source_file.get()
        if not source or not os.path.exists(source):
            messagebox.showerror("Lỗi", "File nguồn không tồn tại!")
            return

        # Get Filters
        filters = {}
        for ent_col, ent_val in self.filter_rows:
            col = ent_col.get().strip()
            val = ent_val.get().strip()
            if col and val:
                filters[col] = [x.strip() for x in val.split(",") if x.strip()]

        mode = self.mode.get()
        split_col = self.entry_split_col.get().strip()
        
        if mode == 2 and not split_col:
            messagebox.showerror("Lỗi", "Vui lòng nhập tên cột để tách file!")
            return

        try:
            self.log("Đang khởi động Excel...")
            app = xw.App(visible=False)
            wb = app.books.open(source)
            sheet = wb.sheets[0] # Assume sheet 1
            
            # Reset existing filters
            if sheet.api.AutoFilterMode:
                sheet.api.AutoFilterMode = False
            
            # Find Headers
            headers = sheet.range("A1").expand("right").value
            header_map = {str(h).strip(): i+1 for i, h in enumerate(headers)} # name -> column index (1-based)

            # Apply Filters
            rng = sheet.range("A1").expand()
            
            for col_name, values in filters.items():
                if col_name in header_map:
                    # xlwings autofilter: criteria1 takes list
                    rng.api.AutoFilter(Field=header_map[col_name], Criteria1=values, Operator=7) # 7=xlFilterValues
                else:
                    print(f"Warning: Column {col_name} not found.")

            # Output folder
            out_dir = os.path.join(os.path.dirname(source), "Output_Exe")
            if not os.path.exists(out_dir):
                os.makedirs(out_dir)

            if mode == 1:
                # MERGE
                self.log("Đang xuất file gộp...")
                new_wb = app.books.add()
                sheet.api.Cells.SpecialCells(12).Copy() # 12 = xlCellTypeVisible
                new_wb.sheets[0].api.Range("A1").PasteSpecial(-4104) # xlPasteAll
                new_wb.sheets[0].api.Range("A1").PasteSpecial(8)     # xlPasteColumnWidths
                
                out_path = os.path.join(out_dir, "Result_Merged.xlsx")
                new_wb.save(out_path)
                new_wb.close()
                messagebox.showinfo("Thành công", f"Đã xuất file tại:\n{out_path}")

            else:
                # SPLIT
                self.log("Đang phân tích dữ liệu tách...")
                if split_col not in header_map:
                    messagebox.showerror("Lỗi", f"Không tìm thấy cột tách '{split_col}'")
                    wb.close()
                    app.quit()
                    return

                col_idx = header_map[split_col]
                
                # Get unique visible values in split column
                # This is tricky with AutoFilter active. 
                # Best way: Copy visible split column to temp sheet, get unique.
                temp_sheet = wb.sheets.add()
                sheet.range(sheet.cells(2, col_idx), sheet.cells(rng.last_cell.row, col_idx)).api.SpecialCells(12).Copy()
                temp_sheet.range("A1").api.PasteSpecial(-4104)
                
                # Get unique values from temp sheet
                vals = temp_sheet.range("A1").expand().value
                if vals is None: vals = []
                if not isinstance(vals, list): vals = [vals]
                unique_vals = list(set([str(v) for v in vals if v is not None]))
                
                temp_sheet.delete()

                self.log(f"Tìm thấy {len(unique_vals)} nhóm. Đang tách...")
                
                for val in unique_vals:
                    # Refilter for this value
                    # We need to keep original filters AND add this one.
                    # Re-apply all filters
                    sheet.api.AutoFilterMode = False
                    rng.api.AutoFilter(Field=col_idx, Criteria1=val)
                    
                    # Re-apply other user filters (must be consistent)
                    for col_name, values in filters.items():
                         if col_name != split_col and col_name in header_map:
                             rng.api.AutoFilter(Field=header_map[col_name], Criteria1=values, Operator=7)

                    # Copy
                    new_wb = app.books.add()
                    sheet.api.Cells.SpecialCells(12).Copy()
                    new_wb.sheets[0].api.Range("A1").PasteSpecial(-4104)
                    new_wb.sheets[0].api.Range("A1").PasteSpecial(8)
                    
                    safe_name = "".join([c if c.isalnum() else "_" for c in val])
                    out_path = os.path.join(out_dir, f"{safe_name}.xlsx")
                    new_wb.save(out_path)
                    new_wb.close()
                
                messagebox.showinfo("Thành công", f"Đã tách {len(unique_vals)} file tại:\n{out_dir}")

        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
            if 'wb' in locals(): wb.close()
        finally:
            if 'app' in locals():
                wb.close()
                app.quit()
            self.log("Hoàn tất.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()
