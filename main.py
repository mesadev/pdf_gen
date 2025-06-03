# main.py
from fastapi import FastAPI, Form, Request
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter
import textwrap
import pandas as pd
import os
import re

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# === CONFIG ===
TEMPLATE_01 = "template/คสศ01.pdf"
TEMPLATE_02 = "template/คสศ02.pdf"
TEMPLATE_03 = "template/คสศ03.pdf"
TEMPLATE_04 = "template/คสศ04.pdf"
TEMPLATE_05 = "template/คสศ05.pdf"
TEMPLATE_06 = "template/คสศ06.pdf"
FONT_PATH = "fonts/THSarabunNew.ttf"
FONT_BOLD_PATH = "fonts/THSarabunNew Bold.ttf"
TEMPLATE_DATA_PATH = "data_project_template/data_pdf3.csv"
price_df = pd.read_csv("data_project_template/price2.csv")

pdfmetrics.registerFont(TTFont("THSarabunNew", FONT_PATH))
pdfmetrics.registerFont(TTFont("THSarabunNew-Bold", FONT_BOLD_PATH))

THAI_MONTHS = {
    '1': 'มกราคม', '01': 'มกราคม',
    '2': 'กุมภาพันธ์', '02': 'กุมภาพันธ์',
    '3': 'มีนาคม', '03': 'มีนาคม',
    '4': 'เมษายน', '04': 'เมษายน',
    '5': 'พฤษภาคม', '05': 'พฤษภาคม',
    '6': 'มิถุนายน', '06': 'มิถุนายน',
    '7': 'กรกฎาคม', '07': 'กรกฎาคม',
    '8': 'สิงหาคม', '08': 'สิงหาคม',
    '9': 'กันยายน', '09': 'กันยายน',
    '10': 'ตุลาคม', '11': 'พฤศจิกายน', '12': 'ธันวาคม'
}

def get_thai_month(month_number):
    return THAI_MONTHS.get(str(int(month_number)), '')

def arabic_to_thai_digits(number_str):
    thai_digits = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")
    return str(number_str).translate(thai_digits)

def extract_numbered_items(text):
    import re
    if not isinstance(text, str):
        return []
    pattern = r'\d+\.\s'
    splits = [m.start() for m in re.finditer(pattern, text)]
    items = []
    for i in range(len(splits)):
        start = splits[i]
        end = splits[i + 1] if i + 1 < len(splits) else None
        item = text[start:end].strip()
        item = re.sub(r'^\d+\.\s*', '', item)
        items.append(item.strip())
    return items

@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("form.html", {"request": request})

@app.post("/generate-pdf")
async def generate_pdf(
    name: str = Form(...),
    position: str = Form(...),
    phone: str = Form(...),
    fund_name: str = Form(...),
    moo: str = Form(...),
    subdistrict: str = Form(...),
    district: str = Form(...),
    province: str = Form(...),
    soi: str = Form(...),
    village: str = Form(...),
    road: str = Form(...),
    house_number: str = Form(...),
    date: str = Form(...),
    community_date: str = Form(...),
    # report_date: str = Form(...),
    project1: str = Form(...),
    budget1: str = Form(...),
    # location: str = Form(...),
    # meeting_time: str = Form(...),
    # total_people: str = Form(...),
    # total_households: str = Form(...),
    # attending_representatives: str = Form(...),
    # male_representatives: str = Form(...),
    # female_representatives: str = Form(...),
    # total_members: str = Form(...),
    # attending_members: str = Form(...),
    # male_members: str = Form(...),
    # female_members: str = Form(...),
    # total_committee: str = Form(...),
):
    writer = PdfWriter()
    df_template = pd.read_csv(TEMPLATE_DATA_PATH)
    matched_row = df_template[df_template['1ชื่อโครงการ'] == project1]

    if matched_row.empty:
        raise ValueError("ไม่พบข้อมูลโครงการใน template")
    row = matched_row.iloc[0]

    day, month, year = date.split("-")[2], date.split("-")[1], date.split("-")[0]
    com_day, com_month, com_year = community_date.split("-")[2], community_date.split("-")[1], community_date.split("-")[0]


    # Clean columns
    price_df["ชื่อโครงการ"] = price_df["ชื่อโครงการ"].astype(str).str.strip()
    price_df["งบประมาณ"] = price_df["งบประมาณ"].astype(str).str.replace(",", "").str.strip().astype(float)
    price_df["งบภาษาไทย"] = price_df["งบภาษาไทย"].astype(str).str.strip()

    # Clean inputs
    project1_clean = project1.strip()
    budget1_float = float(str(budget1).replace(",", "").strip())

    # Match row
    price_match = price_df[
        (price_df["ชื่อโครงการ"] == project1_clean) &
        (price_df["งบประมาณ"] == budget1_float)
    ]

    # Get amount_text
    if price_match.empty:
        amount_text = "ไม่พบข้อมูล"
    else:
        amount_text = price_match.iloc[0]["งบภาษาไทย"]


    

    # ========== หน้า คสศ01 ==========
    overlay1 = BytesIO()
    c = canvas.Canvas(overlay1, pagesize=A4)
    c.setFont("THSarabunNew", 14)
    c.drawString(345, 691, day)
    c.drawString(400, 691, get_thai_month(month))
    c.drawString(488, 691, str(int(year)+543))
    c.drawString(280, 534.7, fund_name)
    c.drawString(410, 534.7, moo)
    c.drawString(465, 534.7, subdistrict)
    c.drawString(113, 515, district)
    c.drawString(238, 515, province)
    c.drawString(153, 496, arabic_to_thai_digits("{:,}".format(int(float(budget1)))))
    c.drawString(250, 496, amount_text)
    c.drawString(130, 476, com_day)
    c.drawString(190, 476, get_thai_month(com_month))
    c.drawString(272, 476, str(int(com_year)+543))
    c.setFont("THSarabunNew-Bold", 10)
    lines = textwrap.wrap(project1, width=40)
    for i, line in enumerate(lines[:2]):
        c.drawString(210, 465 - i*8, line)
    try:
        c.drawString(420, 458, arabic_to_thai_digits("{:,}".format(int(float(budget1)))))
    except:
        c.drawString(410, 458, '-')
    c.setFont("THSarabunNew", 14)
    c.drawString(300, 222.5, name)
    c.drawString(300, 203, position)
    c.drawString(300, 183, phone)
    c.save()
    overlay1.seek(0)

    base01 = PdfReader(TEMPLATE_01)
    page01 = base01.pages[0]
    page01.merge_page(PdfReader(overlay1).pages[0])
    writer.add_page(page01)
    for i in range(1, len(base01.pages)):
        writer.add_page(base01.pages[i])

    # ========== เพิ่มหน้า คสศ02 ==========
    base02 = PdfReader(TEMPLATE_02)

    # --- หน้า 1 ---
    overlay2p1 = BytesIO()
    c1 = canvas.Canvas(overlay2p1, pagesize=A4)
    c1.setFont("THSarabunNew-Bold", 12)
    c1.drawString(210, 521, fund_name)
    c1.drawString(410, 521, house_number)
    c1.drawString(93, 502, soi)
    c1.drawString(170, 502, village)
    c1.drawString(240, 502, road)
    c1.drawString(410, 502, subdistrict)
    c1.drawString(130, 482, district)
    c1.drawString(280, 482, province)
    c1.drawString(440, 482, phone)
    c1.drawString(315, 462, day)
    c1.drawString(385, 462, get_thai_month(month))
    c1.drawString(480, 462, str(int(year)+543))
    c1.setFont("Helvetica", 11)
    for x, y in [(307,443),(367,443),(111,398),(111,379),(292,379),(352,379),(111,340),(111,320),(111,300),(111,281),(111,261),(111,242),(111,222),(111,203)]:
        c1.drawString(x, y, '✓')
    c1.setFont("THSarabunNew-Bold", 10)
    lines = textwrap.wrap(project1, width=40)
    for i, line in enumerate(lines[:2]):
        c1.drawString(230, 164 - i * 8, line)
    try:
        c1.drawString(420, 156, "{:,}".format(int(float(budget1))))
    except:
        c1.drawString(410, 155, '-')
    c1.save()
    overlay2p1.seek(0)
    base02.pages[0].merge_page(PdfReader(overlay2p1).pages[0])
    writer.add_page(base02.pages[0])

    # --- หน้า 2 ---
    overlay2 = BytesIO()
    c2 = canvas.Canvas(overlay2, pagesize=A4)
    c2.setFont("Helvetica", 11)
    for x, y in [(111,655)]:
        c2.drawString(x, y, '✓')

    tick_map = {
        "1": (111, 260),
        "2": (111, 221),
        "3": (111, 201),
        "4": (111, 162),
        "5": (111, 123),
    }

    # อ่านไฟล์ template
    pdf3_df = pd.read_csv(TEMPLATE_DATA_PATH)
    pdf3_df.columns = pdf3_df.columns.str.strip()

    # เทียบชื่อโครงการ
    matched_purpose_row = pdf3_df[
        pdf3_df["1ชื่อโครงการ"].astype(str).str.strip() == project1.strip()
    ]

    # อ่านค่าวัตถุประสงค์
    purpose = ""
    if not matched_purpose_row.empty:
        raw_purpose = matched_purpose_row.iloc[0]["คสศ 02 วัตถุประสงค์การนำงบประมาณไปดำเนินโครงการ"]
        try:
            purpose = str(int(float(raw_purpose)))  # เช่น 5.0 → '5'
        except:
            purpose = str(raw_purpose).strip()

    # วาด ✓ เฉพาะตำแหน่งที่ตรง
    c2.setFont("Helvetica", 11)
    if purpose in tick_map:
        x, y = tick_map[purpose]
        c2.drawString(x, y, "✓")



    c2.setFont("THSarabunNew-Bold", 10)
    lines = textwrap.wrap(project1, width=100 , break_long_words=False)
    for i, line in enumerate(lines[:2]):
        c2.drawString(130, 700 - i * 8, line)
    c2.setFont("THSarabunNew-Bold", 12)
    try:
        c2.drawString(420, 695, "{:,}".format(int(float(budget1))))
    except:
        c2.drawString(410, 695, '-')
    c2.save()
    overlay2.seek(0)
    base02.pages[1].merge_page(PdfReader(overlay2).pages[0])
    writer.add_page(base02.pages[1])

    # --- หน้า 3 ---
    overlay3 = BytesIO()
    c3 = canvas.Canvas(overlay3, pagesize=A4)
    c3.setFont("THSarabunNew-Bold", 12)
    for idx, key in enumerate(['2.เหตุผลและความจำเป็น','4.เป้าหมายการดำเนินโครงการ','5.วิธีการดำเนินโครงการ']):
        lines = textwrap.wrap(row.get(key, ''), width=100, break_long_words=False)
        for i, line in enumerate(lines[:3]):
            c3.drawString(80, 714 - idx*78 - i*20, line)
    steps = extract_numbered_items(row.get('การบริหารจัดการโครงการ', ''))
    people = extract_numbered_items(row.get('ผู้เกี่ยวข้อง', ''))
    times = extract_numbered_items(row.get('ระยะเวลา', ''))
    c3.setFont("THSarabunNew-Bold", 10)
    for i in range(len(steps)):
        y = 456 - i * 18
        c3.drawString(90, y, str(i+1))
        for j, line in enumerate(textwrap.wrap(steps[i], width=40)[:2]):
            c3.drawString(120, y - j*11, line)
        if i < len(people):
            c3.drawString(320, y, people[i])
        if i < len(times):
            c3.drawString(420, y, times[i])


    def extract_numbered_items_dict(text):
        """
        แยกข้อความแบบ '1. xxx 2. yyy' → {1: 'xxx', 2: 'yyy'}
        """
        pattern = r'(\d+)\.\s*(.*?)(?=(?:\d+\.\s)|$)'
        return {int(m.group(1)): m.group(2).strip() for m in re.finditer(pattern, text)}

    # ดึงแถวที่ match project1 + budget1
    if not price_match.empty:
        row_price = price_match.iloc[0]

        # Raw fields
        raw_items = str(row_price["คสศ2 รายละเอียดการใช้เงิน รายการ"])
        raw_amounts = str(row_price["รายละเอียดการใช้เงิน  จำนวนเงินที่ใช้"])
        raw_duration = str(row_price["รายละเอียดการใช้เงิน ระยะเวลา"])

        # Dict by ลำดับ
        item_dict = extract_numbered_items_dict(raw_items)
        amount_dict = extract_numbered_items_dict(raw_amounts)
        duration_dict = extract_numbered_items_dict(raw_duration)

        c3.setFont("THSarabunNew-Bold", 10)
        y = 220

        # แสดงรายการทีละลำดับ (ยึดจาก item_dict)
        for i in sorted(item_dict.keys()):
            c3.drawString(90, y, str(i))                         # ลำดับ
            c3.drawString(120, y, item_dict.get(i, ""))          # รายการ
            c3.drawString(320, y, amount_dict.get(i, ""))        # จำนวนเงิน
            if i == 1:
                c3.drawString(420, y, duration_dict.get(1, ""))  # ระยะเวลาเฉพาะแถวแรก
            y -= 18

        # รวมทั้งสิ้น (แยกบรรทัด)
        total_amount = amount_dict.get(max(amount_dict.keys(), default=0), "")
        total_duration = duration_dict.get(1, "")
        c3.drawString(320, 110, total_amount)
        c3.drawString(420, 110, total_duration)
        c3.drawString(410, 322, total_amount)
        c3.drawString(410, 302, total_amount)
            
    c3.save()
    overlay3.seek(0)
    base02.pages[2].merge_page(PdfReader(overlay3).pages[0])
    writer.add_page(base02.pages[2])

    

    # --- หน้า 4 ---
    overlay4 = BytesIO()
    c4 = canvas.Canvas(overlay4, pagesize=A4)
    c4.setFont("THSarabunNew-Bold", 12)
    y_map = {
        '8.การดำเนินงานโครงการ': 694,
        '9.การบริหารความเสี่ยง': 597,
        '10.การถือครองกรรมสิทธิ์ในทรัพย์สิน': 499,
        '11. การใช้ประโยชน์': 401,
        '12.ความคุ้มค่า และความยั่งยืนของโครงการ': 323,
    }
    for key, y in y_map.items():
        lines = textwrap.wrap(row.get(key, ''), width=100 , break_long_words=False)
        for i, line in enumerate(lines[:3]):
            c4.drawString(80, y - i*20, line)
    c4.drawString(210, 147, name)
    c4.drawString(204, 128, position)
    c4.drawString(215, 108, day)
    c4.drawString(260, 108, get_thai_month(month))
    c4.drawString(350, 108, str(int(year)+543))
    c4.save()
    overlay4.seek(0)
    base02.pages[3].merge_page(PdfReader(overlay4).pages[0])
    writer.add_page(base02.pages[3])

    for i in range(4, len(base02.pages)):
        writer.add_page(base02.pages[i])

     # === คสศ03 ===
    # report_day, report_month, report_year = report_date.split("-")[2], report_date.split("-")[1], report_date.split("-")[0]

    base03 = PdfReader(TEMPLATE_03)
    overlay3p1 = BytesIO()
    c5 = canvas.Canvas(overlay3p1, pagesize=A4)
    c5.setFont("THSarabunNew-Bold", 12)
    c5.drawString(300, 710, fund_name)
    c5.drawString(117, 687.5, moo)
    c5.drawString(170, 687.5, subdistrict)
    c5.drawString(300, 687.5, district)
    c5.drawString(434, 687.5, province)
    # c5.drawString(180, 665, report_day)
    # c5.drawString(240, 665, get_thai_month(report_month))
    # c5.drawString(370, 665, str(int(report_year)+543))
    # c5.drawString(220, 643, location)
    # c5.drawString(170, 604, meeting_time)
    # c5.drawString(345, 564, total_people)
    # c5.drawString(450, 564, total_households)
    # c5.drawString(470, 545, attending_representatives)
    # c5.drawString(210, 525, male_representatives)
    # c5.drawString(290, 525, female_representatives)
    # try:
    #     percent1 = round(int(attending_representatives)/int(total_households)*100)
    # except:
    #     percent1 = 0
    # c5.drawString(412, 525, str(percent1))
    # c5.drawString(315, 505, total_members)
    # c5.drawString(174, 486, attending_members)
    # c5.drawString(262, 486, male_members)
    # c5.drawString(322, 486, female_members)
    # try:
    #     percent2 = round(int(attending_members)/int(total_members)*100)
    # except:
    #     percent2 = 0
    # c5.drawString(419, 486, str(percent2))
    # c5.drawString(435, 466, total_committee)
  
    c5.drawString(180, 273, project1)
    c5.drawString(200, 253, budget1)

    # วาระที่ 2 - ปัญหาความต้องการ
    c5.setFont("THSarabunNew-Bold", 12)
    issues_text = row.get("คสศ 03/1 วาระที่  2 เสนอปัญหาความต้องการ", "")
    issues = extract_numbered_items(issues_text)
    y_pos = 370
    for i, issue in enumerate(issues):
        wrapped_lines = textwrap.wrap(issue, width=80, break_long_words=False)
        for j, line in enumerate(wrapped_lines):
            c5.drawString(140, y_pos, line)
            y_pos -= 19.5

    c5.save()
    overlay3p1.seek(0)
    base03.pages[0].merge_page(PdfReader(overlay3p1).pages[0])
    writer.add_page(base03.pages[0])

    # --- หน้า 2 ---
    overlay3p2 = BytesIO()
    c6 = canvas.Canvas(overlay3p2, pagesize=A4)

    # วาระที่ 3 - ผลที่คาดว่าจะได้รับ
    c6.setFont("THSarabunNew-Bold", 12)
    y_start = 715
    line_spacing = 20

    expected_result = row.get("ผลที่คาดว่าจะได้รับ", "")
    wrapped_lines = textwrap.wrap(expected_result, width=100, break_long_words=False)

    for i, line in enumerate(wrapped_lines[:3]):  
        y = y_start - i * line_spacing
        c6.drawString(90, y, line)


    c6.drawString(160, 404, project1)

    # วาระที่ 3
    issues_text = row.get("วาระที่3 อื่น ๆ ", "")
    issues = extract_numbered_items(issues_text)
    y_pos = 324
    for i, issue in enumerate(issues):
        wrapped_lines = textwrap.wrap(issue, width=120, break_long_words=False)
        for j, line in enumerate(wrapped_lines):
            c6.drawString(85, y_pos, line)
            y_pos -= 19.5

    c6.save()
    overlay3p2.seek(0)
    base03.pages[1].merge_page(PdfReader(overlay3p2).pages[0])
    writer.add_page(base03.pages[1])
    for i in range(2, len(base03.pages)):
        writer.add_page(base03.pages[i])



    # for page in base04.pages:
    #     writer.add_page(page)

    # คสศ 04
    # === หน้า 1 ===
    base04 = PdfReader(TEMPLATE_04)
    overlay4_p1 = BytesIO()
    c4_1 = canvas.Canvas(overlay4_p1, pagesize=A4)
    c4_1.setFont("THSarabunNew-Bold", 12)
    c4_1.drawString(330, 737, fund_name) 
    c4_1.save()
    overlay4_p1.seek(0)
    base04.pages[0].merge_page(PdfReader(overlay4_p1).pages[0])
    writer.add_page(base04.pages[0])

    # === หน้า 2 ===
    overlay4_p2 = BytesIO()
    c4_2 = canvas.Canvas(overlay4_p2, pagesize=A4)
    c4_2.setFont("THSarabunNew-Bold", 12)
    c4_2.drawString(300, 718, fund_name) 
    c4_2.drawString(120, 695, moo)
    c4_2.drawString(170, 695, subdistrict)
    c4_2.drawString(300, 695, district)
    c4_2.drawString(450, 695, province)
    c4_2.save()
    overlay4_p2.seek(0)
    base04.pages[1].merge_page(PdfReader(overlay4_p2).pages[0])
    writer.add_page(base04.pages[1])

 

       # === ตัวอย่างการใส่ข้อความลงบน คสศ05 ===
    base05 = PdfReader(TEMPLATE_05)
    overlay5 = BytesIO()
    c5_1 = canvas.Canvas(overlay5, pagesize=A4)
    c5_1.setFont("THSarabunNew-Bold", 14)

    def extract_numbered_items_dict(text):
        pattern = r'(\d+)\.\s*(.*?)(?=(?:\d+\.\s)|$)'
        return {int(m.group(1)): m.group(2).strip() for m in re.finditer(pattern, text)}

    if not price_match.empty:
        row_price = price_match.iloc[0]

        # ดึงข้อมูลดิบจากแถวที่ match
        raw_items = str(row_price["คสศ2 รายละเอียดการใช้เงิน รายการ"])
        raw_amounts = str(row_price["รายละเอียดการใช้เงิน  จำนวนเงินที่ใช้"])
        raw_duration = str(row_price["รายละเอียดการใช้เงิน ระยะเวลา"])

        # แปลงเป็น dict
        item_dict = extract_numbered_items_dict(raw_items)
        amount_dict = extract_numbered_items_dict(raw_amounts)
        duration_dict = extract_numbered_items_dict(raw_duration)

        c5_1.setFont("THSarabunNew-Bold", 14)
        y = 668  # เริ่มจากตำแหน่งบน

        for i in sorted(item_dict.keys()):
            c5_1.drawString(90, y, str(i))                          # ลำดับ
            c5_1.drawString(130, y, item_dict.get(i, ""))           # รายการ
            c5_1.drawString(340, y, amount_dict.get(i, ""))         # จำนวนเงิน
            if i == 1:
                c5_1.drawString(430, y, duration_dict.get(1, ""))   # ระยะเวลา
            y -= 18

        # รวมทั้งสิ้น
        total_amount = amount_dict.get(max(amount_dict.keys(), default=0), "")
        total_duration = duration_dict.get(1, "")
        c5_1.drawString(340, 447, total_amount)
        c5_1.drawString(430, 447, total_duration)
    


    c5_1.save()  
    overlay5.seek(0)
    base05.pages[0].merge_page(PdfReader(overlay5).pages[0])
    writer.add_page(base05.pages[0])

    for i in range(1, len(base05.pages)):
        writer.add_page(base05.pages[i])
        
    # === ตัวอย่างการใส่ข้อความลงบน คสศ06 ===
    base06 = PdfReader(TEMPLATE_06)
    overlay6 = BytesIO()
    c6_1 = canvas.Canvas(overlay6, pagesize=A4)
    c6_1.setFont("THSarabunNew-Bold", 12)
    c6_1.drawString(335, 716, fund_name)
    c6_1.drawString(293, 628, fund_name)
    c6_1.setFont("THSarabunNew-Bold", 10)

    lines = textwrap.wrap(project1, width=40)
    for i, line in enumerate(lines[:2]):
        c6_1.drawString(204.0, 501.0 - i * 8, line)
    try:
        c6_1.drawString(410, 492.5, "{:,}".format(int(float(budget1))))
    except:
        c6_1.drawString(410, 155, '-')
    
    c6_1.save()

    overlay6.seek(0)
    base06.pages[0].merge_page(PdfReader(overlay6).pages[0])
    writer.add_page(base06.pages[0])
    for i in range(1, len(base06.pages)):
        writer.add_page(base06.pages[i])

    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return StreamingResponse(output, media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=combined.pdf"})