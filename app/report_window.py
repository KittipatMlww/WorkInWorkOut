import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from tkcalendar import DateEntry
import openpyxl
import os
from app.excel_io import logs_between  # เปลี่ยน path ตามจริง

report_window = None

def show_daily_report(root):
    global report_window
    if report_window is not None and report_window.winfo_exists():
        report_window.lift()
        return

    x = root.winfo_x()
    y = root.winfo_y()
    w = root.winfo_width()

    report_window = tk.Toplevel()
    width, height = 950, 500
    report_window.title("รายงานการเข้า-ออกงาน")

    new_x = x + w + 10
    new_y = y
    report_window.geometry(f"{width}x{height}+{new_x}+{new_y}")

    def on_close():
        global report_window
        report_window.destroy()
        report_window = None

    report_window.protocol("WM_DELETE_WINDOW", on_close)

    tk.Label(report_window, text="รายงานการเข้า-ออกงาน",
             font=("TH Sarabun New", 16, "bold")).pack(pady=(10, 0))

    frm = tk.Frame(report_window); frm.pack(pady=(5, 10))

    tk.Label(frm, text="จาก", font=("TH Sarabun New", 14)).pack(side="left")
    date_from = DateEntry(frm, locale="th", date_pattern="dd/MM/yyyy", width=12)
    date_from.pack(side="left", padx=4)

    tk.Label(frm, text="ถึง", font=("TH Sarabun New", 14)).pack(side="left", padx=(10,0))
    date_to = DateEntry(frm, locale="th", date_pattern="dd/MM/yyyy", width=12)
    date_to.pack(side="left", padx=4)

    cols = ("date","in_time","out_time","fullname","id","dept")
    heads = ["วันที่","เวลาเข้า","เวลาออก","ชื่อสกุล","เลขบัตรประชาชน","แผนก"]

    tree = ttk.Treeview(report_window, columns=cols, show="headings")
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    for c,h in zip(cols,heads):
        tree.heading(c, text=h)
        tree.column(c, anchor="center", width=130)
    tree.column("fullname", anchor="w", width=200)

    ysb = ttk.Scrollbar(report_window, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=ysb.set)
    ysb.pack(side="right", fill="y")

    # ===== 🔁 Export Excel Function =====
    def export_to_excel():
        rows = [tree.item(child)['values'] for child in tree.get_children()]
        if not rows:
            messagebox.showinfo("ไม่มีข้อมูล", "ไม่มีข้อมูลให้ส่งออก")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="บันทึกรายงานเป็น Excel",
            initialfile="รายงานการเข้างาน"  # 🆕 เพิ่มบรรทัดนี้
        )
        if not file_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(heads)  # หัวตาราง
        for row in rows:
            new_row = list(row)
            if len(new_row) >= 5:
                new_row[4] = str(new_row[4]).zfill(13)  # แก้ตรงนี้!
            ws.append(new_row)
        try:
            wb.save(file_path)
            messagebox.showinfo("บันทึกสำเร็จ", "ส่งออกรายงานเรียบร้อยแล้ว")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{e}")

    # ===== 📋 Right-click Menu =====
    context_menu = tk.Menu(report_window, tearoff=0)
    context_menu.add_command(label="ส่งออกเป็น Excel", command=export_to_excel)

    def show_context_menu(event):
        context_menu.post(event.x_root, event.y_root)

    tree.bind("<Button-3>", show_context_menu)

    def clear_tree():
        for i in tree.get_children():
            tree.delete(i)

    def search():
        try:
            d1 = datetime.strptime(date_from.get(), "%d/%m/%Y")
            d2 = datetime.strptime(date_to.get(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("รูปแบบวันที่ผิด", "กรุณาเลือกวันที่ให้ถูกต้อง")
            return
        if d1 > d2:
            messagebox.showwarning("ช่วงวันที่ไม่ถูกต้อง", "วันที่เริ่มต้นต้องไม่มากกว่าวันที่สิ้นสุด")
            return
        clear_tree()
        for rec in logs_between(d1, d2):
            tree.insert("", "end", values=[
                rec["วันที่"], rec["เวลาเข้า"], rec["เวลาออก"],
                rec["ชื่อสกุล"], rec["เลขบัตรประชาชน"], rec["แผนก"]
            ])

    tk.Button(frm, text="ค้นหา", command=search,
              font=("TH Sarabun New", 12), width=8).pack(side="left", padx=10)

    search()

if __name__ == "__main__":
    root = tk.Tk()
    tk.Button(root, text="เปิดรายงาน", command=lambda: show_daily_report(root)).pack(padx=20, pady=20)
    root.mainloop()
