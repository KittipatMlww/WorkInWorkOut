import tkinter as tk
from datetime import datetime
import locale
from app.employee_list_window import show_employee_list
from app.issue_card_window import open_issue_card_window

# ตั้งค่าภาษาไทยสำหรับวันที่ (Linux/Mac ใช้ "th_TH.UTF-8", Windows อาจใช้ "Thai_Thailand")
try:
    locale.setlocale(locale.LC_TIME, "th_TH.UTF-8")
except:
    locale.setlocale(locale.LC_TIME, "Thai_Thailand")

today = datetime.now().strftime("วันที่ %d/%m/%Y")
def run_app():
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
    # ใช้ function แบบไม่ต้องส่ง root ตรงนี้
    # ปุ่ม “รายชื่อพนักงาน” เป็นสีฟ้า-ข้อความขาว
    btn1 = tk.Button(
        root,
        text="รายชื่อพนักงาน",
        width=18,
        height=4,
        # bg="#007bff",      # พื้นหลังสีน้ำเงิน (bootstrap primary)
        # fg="white",        # ตัวหนังสือสีขาว
        activebackground="#0056b3",  # พื้นหลังเมื่อกดเป็นสีน้ำเงินเข้ม
        activeforeground="white",
        command=lambda: show_employee_list(root)
    )
    btn1.grid(row=3, column=0, padx=5, pady=5)

    # ปุ่ม “ออกReport” เป็นสีเขียว (แต่แช่ disabled ไว้)
    btn2 = tk.Button(
        root,
        text="ออกReport",
        width=18,
        height=4,
        # bg="#28a745",      # พื้นหลังสีเขียว (bootstrap success)
        # fg="white",
        disabledforeground="#e0e0e0",  # ตัวหนังสือสีเทาเมื่อ disabled
        state="disabled"
    )
    btn2.grid(row=3, column=1, padx=5, pady=5)

    # ปุ่ม “เริ่มออกบัตร” เป็นสีส้ม (แต่แช่ disabled ไว้)
    btn3 = tk.Button(
        root,
        text="เริ่มออกบัตร",
        width=18,
        height=4,
        # bg="#fd7e14",      # พื้นหลังสีส้ม (bootstrap warning)
        # fg="white",
        activebackground="#0056b3",  # พื้นหลังเมื่อกดเป็นสีน้ำเงินเข้ม
        activeforeground="white",
        command=lambda:open_issue_card_window(root)
    )
    btn3.grid(row=3, column=2, padx=5, pady=5)
    root.update()
    root.minsize(root.winfo_width(), root.winfo_height())
    root.mainloop()
