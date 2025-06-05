import openpyxl # ใช้สำหรับอ่าน/เขียนไฟล์ Excel (.xlsx)
import os # ใช้ตรวจสอบไฟล์และจัดการ path

#กำหนด path ของไฟล์เก็บข้อมูลพนักงานที่อยู่ในโฟลเดอร์ assets/ ชื่อไฟล์ employees.xlsx
EMPLOYEE_FILE = "assets/employees.xlsx" 

def load_employees():
    if not os.path.exists(EMPLOYEE_FILE): #ตรวจสอบว่าไฟล์ employees.xlsx มีอยู่จริงหรือไม่
        return [] #ถ้าไม่มีคืนค่าลิสต์ว่าง
    wb = openpyxl.load_workbook(EMPLOYEE_FILE)
    ws = wb.active
    return [(str(row[0]), row[1], row[2]) for row in ws.iter_rows(min_row=2, values_only=True)]

def save_employees(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["เลขบัตรประชาชน", "ชื่อ-สกุล", "แผนก"])
    for row in data:
        ws.append([str(row[0]), row[1], row[2]])
    os.makedirs(os.path.dirname(EMPLOYEE_FILE), exist_ok=True)
    wb.save(EMPLOYEE_FILE)