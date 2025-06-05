import tkinter as tk
from datetime import datetime, date
import locale
from app.excel_io import load_employees, save_employees  # ✅ โหลดจากไฟล์ Excel

# 🔒 เก็บสถานะว่าเปิดหน้าต่างแล้วหรือยัง
issue_card_window = None

try:
    locale.setlocale(locale.LC_TIME, "th_TH.UTF-8")
except:
    locale.setlocale(locale.LC_TIME, "Thai_Thailand")

# ⏱️ บันทึกการเข้าออกชั่วคราว (รีใหม่ทุกวัน)
daily_time_log = {}

def open_issue_card_window(root):
    global issue_card_window
    # ✋ ป้องกันเปิดซ้ำ
    if issue_card_window is not None and issue_card_window.winfo_exists():
        issue_card_window.lift()  # ดึงขึ้นมาด้านหน้า
        return
    # ดึงตำแหน่งของหน้าต่างหลัก
    x = root.winfo_x()
    y = root.winfo_y()
    w = root.winfo_width()

    issue_card_window = tk.Toplevel()
    # กำหนดขนาด
    width, height = 550, 350
    issue_card_window.title("เริ่มออกบัตรพนักงาน")
    # ตำแหน่งใหม่: ด้านขวาของหน้าต่างหลัก
    new_x = x + w + 10  # +10 เผื่อระยะห่าง
    new_y = y
    issue_card_window.geometry(f"{width}x{height}+{new_x}+{new_y}")

    def on_close():
        global issue_card_window
        issue_card_window.destroy()
        issue_card_window = None  # 💡 คืนค่าเมื่อปิดหน้าต่าง

    issue_card_window.protocol("WM_DELETE_WINDOW", on_close)


    time_label = tk.Label(issue_card_window, text="", font=("Arial", 48))
    time_label.grid(row=0, column=0, columnspan=2, pady=(10, 5))

    def update_time():
        now = datetime.now().strftime("%H:%M:%S")
        time_label.config(text=now)
        issue_card_window.after(1000, update_time)

    update_time()

    date_str = datetime.now().strftime("วันที่ %d/%m/%Y")
    date_label = tk.Label(issue_card_window, text=date_str, font=("TH Sarabun New", 14))
    date_label.grid(row=1, column=1, sticky="e", padx=20, pady=(0, 10))

    fields = {
        "ชื่อสกุล": tk.StringVar(),
        "แผนก": tk.StringVar(),
        "เวลาเข้า": tk.StringVar(),
        "เวลาออก": tk.StringVar(),
    }

    row_index = 2
    for label, var in fields.items():
        tk.Label(issue_card_window, text=label, font=("TH Sarabun New", 14)).grid(row=row_index, column=0, sticky="e", padx=10, pady=3)
        entry = tk.Entry(issue_card_window, textvariable=var, font=("TH Sarabun New", 14), state="readonly", width=25)
        entry.grid(row=row_index, column=1, sticky="w", padx=10)
        row_index += 1

    # 📥 โหลดข้อมูลพนักงานจาก Excel
    employee_data = load_employees()
    # 💡 แปลงเป็น dict โดยใช้เลขบัตรเป็น key
    employee_dict = {row[0]: {"fullname": row[1], "dept": row[2]} for row in employee_data}

    # 🕒 เก็บวันที่ที่โหลดล่าสุดไว้เพื่อ reset ข้อมูลเมื่อข้ามวัน
    last_checked_date = [date.today()]

    def reset_daily_log_if_new_day():
        today = date.today()
        if today != last_checked_date[0]:
            daily_time_log.clear()
            last_checked_date[0] = today

    def handle_code_enter(event=None):
        reset_daily_log_if_new_day()

        code = entry_code.get().strip()
        if code in employee_dict:
            now_str = datetime.now().strftime("%H:%M:%S")
            info = employee_dict[code]

            fields["ชื่อสกุล"].set(info["fullname"])
            fields["แผนก"].set(info["dept"])

            if code not in daily_time_log:
                daily_time_log[code] = {"in": now_str, "out": ""}
                fields["เวลาเข้า"].set(now_str)
                fields["เวลาออก"].set("")
            elif daily_time_log[code]["out"] == "":
                daily_time_log[code]["out"] = now_str
                fields["เวลาเข้า"].set(daily_time_log[code]["in"])
                fields["เวลาออก"].set(now_str)
            else:
                fields["เวลาเข้า"].set(daily_time_log[code]["in"])
                fields["เวลาออก"].set(daily_time_log[code]["out"])
        else:
            for var in fields.values():
                var.set("")

    tk.Label(issue_card_window, text="โปรดใส่เลขบัตรประชาชน", font=("TH Sarabun New", 14)).grid(row=row_index, column=0, sticky="e", padx=10, pady=(10, 5))
    entry_code = tk.Entry(issue_card_window, font=("TH Sarabun New", 14), width=25)
    entry_code.grid(row=row_index, column=1, sticky="w", padx=10, pady=(10, 5))
    entry_code.bind("<Return>", handle_code_enter)
