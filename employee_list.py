import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import os

EMPLOYEE_FILE = "employees.xlsx"

def load_employees():
    if not os.path.exists(EMPLOYEE_FILE):
        return []
    wb = openpyxl.load_workbook(EMPLOYEE_FILE)
    ws = wb.active
    return [(str(row[0]), row[1], row[2]) for row in ws.iter_rows(min_row=2, values_only=True)]

def save_employees(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["เลขบัตรประชาชน", "ชื่อ-สกุล", "แผนก"])
    for row in data:
        ws.append([str(row[0]), row[1], row[2]])
    wb.save(EMPLOYEE_FILE)

def show_employee_list(root):
    window = tk.Toplevel(root)
    window.title("รายชื่อพนักงาน")
    window.geometry("600x400")
    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)

    employee_data = load_employees()

    def add_or_edit():
        cid = entry_id.get().strip()
        name = entry_name.get().strip()
        dept = entry_dept.get().strip()
        if not cid or not name or not dept:
            messagebox.showwarning("กรอกข้อมูลไม่ครบ", "กรุณากรอกให้ครบ")
            return
        if not cid.isdigit() or len(cid) != 13:
            messagebox.showerror("เลขบัตรไม่ถูกต้อง", "เลขบัตรประชาชนต้องเป็นตัวเลข 13 หลัก")
            return
        
        # ตรวจสอบชื่อซ้ำ ยกเว้นกรณีที่ชื่อเดิมตรงกับ cid นี้
        for row in employee_data:
            existing_cid, existing_name, _ = row
            if existing_cid != cid and existing_cid == cid:
                messagebox.showerror("เลขบัตรซ้ำ", f"เลขบัตรประชาชน '{cid}' มีอยู่ในระบบแล้ว")
                return
            if existing_cid != cid and existing_name == name:
                messagebox.showerror("ชื่อซ้ำ", f"ชื่อ '{name}' มีอยู่ในระบบแล้ว")
                return

        for i, row in enumerate(employee_data):
            if row[0] == cid:
                confirm = messagebox.askyesno("ยืนยันการแก้ไข", f"มีเลขบัตรประชาชน '{cid}' อยู่แล้ว\nคุณต้องการแก้ไขข้อมูลนี้หรือไม่?")
                if confirm:
                    employee_data[i] = (cid, name, dept)
                    update_table()
                    save_employees(employee_data)
                    clear_inputs()
                else:
                    messagebox.showinfo("ยกเลิก", "ไม่ได้มีการเปลี่ยนแปลงข้อมูล")
                    clear_inputs()
                return  # จบฟังก์ชันหลังจากตอบแล้ว

        # ถ้าไม่เจอ -> เพิ่มใหม่
        employee_data.append((cid, name, dept))
        update_table()
        save_employees(employee_data)
        clear_inputs()

    def delete_selected():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("ไม่ได้เลือก", "กรุณาเลือกแถวก่อนลบ")
            return
        cid = str(tree.item(selected)["values"][0])
        nonlocal employee_data
        employee_data = [row for row in employee_data if str(row[0]) != cid]
        update_table()
        save_employees(employee_data)
        clear_inputs()

    def update_table():
        tree.delete(*tree.get_children())
        for row in employee_data:
            tree.insert("", "end", values=row)

    def clear_inputs():
        entry_id.delete(0, tk.END)
        entry_name.delete(0, tk.END)
        entry_dept.delete(0, tk.END)

    def on_row_select(event):
        selected = tree.selection()
        if selected:
            values = tree.item(selected)["values"]
            entry_id.delete(0, tk.END)
            entry_id.insert(0, values[0])
            entry_name.delete(0, tk.END)
            entry_name.insert(0, values[1])
            entry_dept.delete(0, tk.END)
            entry_dept.insert(0, values[2])

    # UI layout
    main_frame = tk.Frame(window)
    main_frame.grid(row=0, column=0, sticky="nsew")
    main_frame.grid_rowconfigure(0, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)

    tree = ttk.Treeview(main_frame, columns=("cid", "name", "dept"), show="headings", selectmode="browse")
    tree.heading("cid", text="เลขบัตรประชาชน")
    tree.heading("name", text="ชื่อ-สกุล")
    tree.heading("dept", text="แผนก")
    tree.grid(row=0, column=0, columnspan=5, sticky="nsew", padx=5, pady=5)
    tree.bind("<<TreeviewSelect>>", on_row_select)

    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=5, sticky='ns')

    entry_id = tk.Entry(main_frame, width=20)
    entry_id.grid(row=1, column=0, padx=5, pady=5)

    entry_name = tk.Entry(main_frame, width=25)
    entry_name.grid(row=1, column=1, padx=5, pady=5)

    entry_dept = tk.Entry(main_frame, width=15)
    entry_dept.grid(row=1, column=2, padx=5, pady=5)

    btn_add = tk.Button(main_frame, text="Add/Edit", command=add_or_edit)
    btn_add.grid(row=1, column=3, padx=5, pady=5)

    btn_del = tk.Button(main_frame, text="Del", command=delete_selected)
    btn_del.grid(row=1, column=4, padx=5, pady=5)

    update_table()
