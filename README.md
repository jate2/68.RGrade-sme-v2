# Student Grade Report Generator (SME Version)

Streamlit app สำหรับสร้างรายงานผลการเรียนรายบุคคลจากไฟล์ Excel และ Word Template

## ฟีเจอร์หลัก
- อัปโหลด Excel ไฟล์เกรดนักเรียน
- ใช้ Word Template ที่มี Placeholder
- เลือกรายงานรายคน/หลายคน
- ดาวน์โหลดเป็น .docx หรือ .pdf
- แสดง Dashboard preview ข้อมูลนักเรียน

## วิธีใช้งาน

```bash
pip install -r requirements.txt
streamlit run report_generator_web.py
```

## Deploy บน Streamlit Cloud

1. อัปโหลดไฟล์ทั้งหมดไปที่ GitHub Repo
2. สร้างแอปใหม่บน [https://streamlit.io/cloud](https://streamlit.io/cloud)
3. เลือก repo และไฟล์ `report_generator_web.py`