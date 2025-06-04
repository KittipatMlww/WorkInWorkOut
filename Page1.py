import tkinter as tk
from datetime import datetime
import locale
import employee_list 
# ตั้งค่าภาษาไทยสำหรับวันที่ (Linux/Mac ใช้ "th_TH.UTF-8", Windows อาจใช้ "Thai_Thailand")
try:
    locale.setlocale(locale.LC_TIME, "th_TH.UTF-8")
except:
    locale.setlocale(locale.LC_TIME, "Thai_Thailand")

today = datetime.now().strftime("วันที่ %d/%m/%Y")

# # ฟังก์ชันจำลอง
# def export_report():
#     messagebox.showinfo("ไปหน้า", "ออก Report")

# def start_card_process():
#     messagebox.showinfo("ไปหน้า", "เริ่มออกบัตร")

# สร้างหน้าต่างหลัก
root = tk.Tk()
root.title("โปรแกรมลงบัตรพนักงาน")

company_name = "Kimleng (Thailand) Public Company Limited"

# ========== แถวที่ 0: ชื่อบริษัท ==========
lbl_title = tk.Label(root, text=company_name, font=("TH Sarabun New", 16, "bold"))
lbl_title.grid(row=0, column=0, padx=5, columnspan=3, pady=(10, 5))

# ========== แถวที่ 1: วันที่ และ ปุ่ม เริ่มต้น ==========
lbl_date = tk.Label(root, text=today, font=("TH Sarabun New", 12))
lbl_date.grid(row=1, column=0, padx=10, pady=(10, 15), sticky="w")

btn_init = tk.Button(root, text="ทำการเริ่มต้น", state="disabled", width=20)
btn_init.grid(row=1, column=1, padx=5, pady=(10, 15), sticky="e")

# ========== แถวที่ 3: ปุ่ม 3 ปุ่มหลัก ==========
btn1 = tk.Button(root, text="รายชื่อพนักงาน", width=18, height=4, command=lambda: employee_list.show_employee_list(root))
btn1.grid(row=3, column=0, padx=5, pady=5)

btn2 = tk.Button(root, text="ออกReport", width=18, height=4, state="disabled")
btn2.grid(row=3, column=1, padx=5, pady=5)

btn3 = tk.Button(root, text="เริ่มออกบัตร", width=18, height=4, state="disabled")
btn3.grid(row=3, column=2, padx=5, pady=5)


# คำสั่งนี้ทำให้หน้าต่างพอดีอัตโนมัติ
root.update()
root.minsize(root.winfo_width(), root.winfo_height())

# เริ่มแสดง GUI
root.mainloop()
