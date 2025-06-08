from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import date

# ✅ โหลดฟอนต์ที่คุณมีอยู่จริง
pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabun.ttf'))

def create_medical_form(name, cid, dob, address, filename="ใบรับรองแพทย์.pdf"):
    c = canvas.Canvas(filename)
    c.setFont("THSarabun", 18)

    c.drawCentredString(300, 800, "ใบรับรองแพทย์")
    c.drawString(100, 750, f"ชื่อ-นามสกุล: {name}")
    c.drawString(100, 730, f"เลขบัตรประชาชน: {cid}")
    c.drawString(100, 710, f"วันเกิด: {dob}")
    c.drawString(100, 690, f"ที่อยู่: {address}")
    c.drawString(100, 650, f"วันที่ออกใบรับรอง: {date.today().strftime('%d/%m/%Y')}")
    c.drawString(100, 610, "ผลการตรวจ: ___________________________")
    c.drawString(100, 590, "แพทย์ผู้ตรวจ: __________________________")
    
    c.showPage()
    c.save()
    print("✅ สร้างใบรับรองแพทย์เรียบร้อย:", filename)
