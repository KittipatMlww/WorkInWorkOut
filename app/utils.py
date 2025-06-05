#นำเข้าโมดูลมาตรฐานของ Python สำหรับการจัดการกับ Regular Expressions (regex) 
#ใช้ในการค้นหา ตรวจสอบ หรือจัดการข้อความ
import re

#ตรวจว่าเป็นตัวเลขล้วนและต้อง 13 หลัก
def is_valid_citizen_id(cid):
    return cid.isdigit() and len(cid) == 13

#ตรวจว่าเป็นภาษาไทยล้วนหรือภาษาอังกฤษล้วนและช่องว่างเท่านั้น
def is_valid_name(name):
    return bool(re.match(r"^[\u0E00-\u0E7Fa-zA-Z\s]+$", name))