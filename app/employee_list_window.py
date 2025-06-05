import tkinter as tk
from tkinter import ttk, messagebox
from app.excel_io import load_employees, save_employees
from app.utils import is_valid_citizen_id, is_valid_name

# 🔒 เก็บสถานะว่าเปิดหน้าต่างแล้วหรือยัง
employee_list_window = None

def show_employee_list(root):
    global employee_list_window

    # ✋ ป้องกันเปิดซ้ำ
    if employee_list_window is not None and employee_list_window.winfo_exists():
        employee_list_window.lift()  # ดึงขึ้นมาด้านหน้า
        return
    # ดึงตำแหน่งของหน้าต่างหลัก
    x = root.winfo_x()
    y = root.winfo_y()
    w = root.winfo_width()

    employee_list_window = tk.Toplevel()
    employee_list_window.title("รายชื่อพนักงาน")
    # กำหนดขนาด
    width, height = 700, 450
    # ตำแหน่งใหม่: ด้านขวาของหน้าต่างหลัก
    new_x = x + w + 10  # +10 เผื่อระยะห่าง
    new_y = y
    employee_list_window.geometry(f"{width}x{height}+{new_x}+{new_y}")

    def on_close():
        global employee_list_window
        employee_list_window.destroy()
        employee_list_window = None  # 💡 คืนค่าเมื่อปิดหน้าต่าง

    employee_list_window.protocol("WM_DELETE_WINDOW", on_close)

    employee_data = load_employees()  # 📥 โหลดข้อมูลพนักงานจากไฟล์

    def add_or_edit():
        cid = entry_id.get().strip()
        name = entry_name.get().strip()
        dept = entry_dept.get().strip()

        # ❗ ตรวจสอบว่ากรอกข้อมูลครบ
        if not cid or not name or not dept:
            messagebox.showwarning("กรอกข้อมูลไม่ครบ", "กรุณากรอกให้ครบ")
            return
        if not is_valid_citizen_id(cid):
            messagebox.showerror("เลขบัตรไม่ถูกต้อง", "เลขบัตรประชาชนต้องเป็นตัวเลข 13 หลัก")
            return
        if not is_valid_name(name):
            messagebox.showerror("ชื่อไม่ถูกต้อง", "ชื่อควรมีแต่ตัวอักษรไทยหรืออังกฤษเท่านั้น")
            return

        # ❌ ป้องกันชื่อซ้ำกับคนอื่น (ถ้าเลขบัตรไม่ตรง)
        for row in employee_data:
            existing_cid, existing_name, _ = row
            if existing_cid != cid and existing_name == name:
                messagebox.showerror("ชื่อซ้ำ", f"ชื่อ '{name}' มีอยู่ในระบบแล้ว")
                return

        # 🔄 แก้ไขข้อมูลถ้าเจอเลขบัตรซ้ำ
        for i, row in enumerate(employee_data):
            if row[0] == cid:
                confirm = messagebox.askyesno(
                    "ยืนยันการแก้ไข",
                    f"มีเลขบัตรประชาชน '{cid}' อยู่แล้ว\nคุณต้องการแก้ไขข้อมูลนี้หรือไม่?"
                )
                if confirm:
                    employee_data[i] = (cid, name, dept)
                    update_table()
                    save_employees(employee_data)
                    messagebox.showinfo("สำเร็จ", "แก้ไขพนักงานเรียบร้อยแล้ว")
                clear_inputs()
                return

        # ➕ เพิ่มพนักงานใหม่
        employee_data.append((cid, name, dept))
        update_table()
        save_employees(employee_data)
        messagebox.showinfo("สำเร็จ", "เพิ่มพนักงานเรียบร้อยแล้ว")

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
        messagebox.showinfo("ลบแล้ว", "ลบพนักงานเรียบร้อยแล้ว")

    def update_table():
        tree.delete(*tree.get_children())
        for index, row in enumerate(employee_data):
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            tree.insert("", "end", values=row, tags=(tag,))
        clear_inputs()

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

    # 📦 สร้าง frame หลัก
    frame = tk.Frame(employee_list_window, padx=10, pady=10, relief="solid", borderwidth=1)
    frame.pack(fill=tk.BOTH, expand=True)

    # 📋 ตารางรายชื่อพนักงาน
    tree = ttk.Treeview(
        frame,
        columns=("cid", "name", "dept"),
        show="headings",
        height=12
    )
    tree.heading("cid", text="เลขบัตรประชาชน")
    tree.heading("name", text="ชื่อ-สกุล")
    tree.heading("dept", text="แผนก")
    tree.column("cid", anchor="center", width=200)
    tree.column("name", anchor="center", width=250)
    tree.column("dept", anchor="center", width=100)

    tree.tag_configure("oddrow", background="#f2f2f2")
    tree.tag_configure("evenrow", background="#ffffff")

    tree.grid(row=0, column=0, columnspan=5, sticky="nsew")
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    tree.bind("<<TreeviewSelect>>", on_row_select)

    # 📜 แถบเลื่อน
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=5, sticky='ns')

    # 🧾 ช่องกรอกข้อมูล
    entry_id = ttk.Entry(frame, width=20)
    entry_id.grid(row=1, column=0, padx=5, pady=10)

    entry_name = ttk.Entry(frame, width=25)
    entry_name.grid(row=1, column=1, padx=5, pady=10)

    entry_dept = ttk.Entry(frame, width=15)
    entry_dept.grid(row=1, column=2, padx=5, pady=10)

    # 🔘 ปุ่มเพิ่ม/แก้ไข
    btn_add = ttk.Button(frame, text="Add/Edit", command=add_or_edit)
    btn_add.grid(row=1, column=3, padx=5, pady=10)

    # ❌ ปุ่มลบ
    btn_del = ttk.Button(frame, text="Del", command=delete_selected)
    btn_del.grid(row=1, column=4, padx=5, pady=10)

    update_table()
