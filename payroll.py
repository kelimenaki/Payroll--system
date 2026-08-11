import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os
import sys
import webbrowser
import urllib.parse

# ====== COMPANY DETAILS ======
COMPANY_NAME = "YOUR COMPANY LTD"
COMPANY_ADDRESS = "P.O Box 12345-00100, Nairobi, Kenya"
COMPANY_KRA_PIN = "A001234567X"
COMPANY_PHONE = "+254 700 123456"
LOGO_PATH = "logo.png"
EXCEL_FILE = "Kenya_Payroll_Full_2026.xlsx"
# =================================================

# ====== GET PREVIOUS MONTH AUTOMATICALLY ======
prev_month_date = datetime.now().replace(day=1) - pd.DateOffset(months=1)
PAYROLL_MONTH = prev_month_date.strftime("%B %Y") # e.g. "July 2026"
PAYROLL_MONTH_SHORT = prev_month_date.strftime("%b%Y") # e.g. "Jul2026"
# =================================================

# ====== KRA RATES 2026 ======
NSSF_RATE = 0.06
NSSF_CAP = 6480
SHIF_RATE = 0.0275
HOUSING_RATE = 0.015
PAYE_RELIEF = 2400

PAYE_BANDS = [
    [24000, 0.10, 0],
    [32333, 0.25, 2400],
    [500000, 0.30, 4483.25],
    [800000, 0.325, 145083.25],
    [999999999, 0.35, 242583.25]
]
# ========================================================

def create_new_file():
    print("Creating new payroll file...")
    names = ["Achieng Otieno","Brian Kamau","Cynthia Wanjiku","David Mwangi","Esther Akinyi"]
    gross = [25000,30000,35000,40000,50000]
    pension = [0,0,1000,0,500]
    helb = [0,1000,0,2500,0]
    sha_numbers = ["CR0MGS4561569-1","CR0MGS4561570-1","CR0MGS4561571-1","CR0MGS4561572-1","CR0MGS4561573-1"]
    phones = [f"07{10000000+i}" for i in range(5)]
    data = []
    for i in range(5):
        data.append([names[i], f"{10000000+i+1}", f"A00{100000+i+1}X", sha_numbers[i], phones[i], gross[i], pension[i], helb[i]])
    df = pd.DataFrame(data, columns=["Employee Name","ID Number","KRA PIN","SHA Number","Phone Number","Gross Salary","Pension","HELB"])
    return df

def load_from_excel():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ {EXCEL_FILE} not found!")
        sys.exit()
    print(f"Reading {EXCEL_FILE}...")
    return pd.read_excel(EXCEL_FILE, sheet_name='Payroll', engine='openpyxl')

def add_formulas(df):
    for i in range(2, len(df)+2):
        df.loc[i-2, 'NSSF'] = f'=MIN(F{i}*{NSSF_RATE},{NSSF_CAP})'
        df.loc[i-2, 'SHIF'] = f'=F{i}*{SHIF_RATE}'
        df.loc[i-2, 'Housing Levy'] = f'=F{i}*{HOUSING_RATE}'
        df.loc[i-2, 'Taxable Pay'] = f'=F{i}-G{i}-H{i}-I{i}-J{i}'

        paye_if = f'IF(L{i}<={PAYE_BANDS[0][0]},L{i}*{PAYE_BANDS[0][1]}'
        for band in PAYE_BANDS[1:]:
            prev = PAYE_BANDS[PAYE_BANDS.index(band)-1][0]
            paye_if += f',IF(L{i}<={band[0]},{band[2]}+(L{i}-{prev})*{band[1]}'
        paye_if += ')' * len(PAYE_BANDS)
        df.loc[i-2, 'PAYE'] = f'=MAX(0,ROUND({paye_if}-{PAYE_RELIEF},0))'
        df.loc[i-2, 'Net Pay'] = f'=F{i}-G{i}-H{i}-I{i}-J{i}-L{i}'
    return df

def calc_paye(taxable):
    paye_raw = 0
    prev_limit = 0
    for limit, rate, base in PAYE_BANDS:
        if taxable <= limit:
            paye_raw = base + (taxable - prev_limit) * rate
            break
        prev_limit = limit
    return max(0, round(paye_raw - PAYE_RELIEF, 0))

def build_other_tabs(df):
    summary_data = [
        ['GROSS PAY',f'=SUM(Payroll!F2:F{len(df)+1})',''],
        ['NSSF',f'=SUM(Payroll!I2:I{len(df)+1})',f'=SUM(Payroll!I2:I{len(df)+1})'],
        ['SHIF',f'=SUM(Payroll!J2:J{len(df)+1})',''],
        ['HOUSING LEVY',f'=SUM(Payroll!K2:K{len(df)+1})',f'=SUM(Payroll!K2:K{len(df)+1})'],
        ['PAYE',f'=SUM(Payroll!M2:M{len(df)+1})',''],
        ['NET PAY',f'=SUM(Payroll!N2:N{len(df)+1})',''],
    ]
    summary = pd.DataFrame(summary_data, columns=['Deduction','Employee Total','Employer Total/Notes'])
    phones_tab = pd.DataFrame()
    phones_tab['Employee Name'] = df['Employee Name']
    phones_tab['Phone Number'] = df['Phone Number']
    phones_tab['Net Pay'] = df['Net Pay']
    cleartax = pd.DataFrame()
    cleartax['Employee Name'] = df['Employee Name']
    cleartax['ID Number'] = df['ID Number']
    cleartax['KRA PIN'] = df['KRA PIN']
    cleartax['SHA Number'] = df['SHA Number']
    cleartax['Gross Pay'] = df['Gross Salary']
    cleartax['NSSF Deduction'] = df['NSSF']
    cleartax['SHIF Deduction'] = df['SHIF']
    cleartax['Housing Levy Deduction'] = df['Housing Levy']
    cleartax['PAYE'] = df['PAYE']
    cleartax['Net Pay'] = df['Net Pay']
    return summary, phones_tab, cleartax

def generate_pdfs(df):
    os.makedirs("Payslips", exist_ok=True)
    print(f"Generating payslips for: {PAYROLL_MONTH}")
    for i in range(len(df)):
        gross = float(df.loc[i,'Gross Salary'])
        pension = float(df.loc[i,'Pension'])
        helb = float(df.loc[i,'HELB'])
        sha_no = df.loc[i,'SHA Number']
        nssf = min(gross * NSSF_RATE, NSSF_CAP)
        shif = gross * SHIF_RATE
        housing = gross * HOUSING_RATE
        taxable = gross - pension - helb - nssf - shif
        paye = calc_paye(taxable)
        net_pay = gross - pension - helb - nssf - shif - housing - paye

        filename = f"Payslips/Payslip_{df.loc[i,'Employee Name'].replace(' ','_')}_{PAYROLL_MONTH_SHORT}.pdf"
        c = canvas.Canvas(filename, pagesize=A4)
        width, height = A4
        try:
            logo = ImageReader(LOGO_PATH)
            c.drawImage(logo, 50, height-90, width=60, height=60, preserveAspectRatio=True)
        except: pass
        c.setFont("Helvetica-Bold", 14)
        c.drawString(120, height-60, COMPANY_NAME)
        c.setFont("Helvetica", 9)
        c.drawString(120, height-75, COMPANY_ADDRESS)
        c.drawString(120, height-88, f"KRA PIN: {COMPANY_KRA_PIN} | Tel: {COMPANY_PHONE}")
        c.line(50, height-100, width-50, height-100)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(width/2, height-120, "PAYSLIP")
        c.setFont("Helvetica", 10)
        c.drawCentredString(width/2, height-135, f"For the period: {PAYROLL_MONTH}")
        y = height - 165
        c.drawString(50, y, f"Employee Name: {df.loc[i,'Employee Name']}")
        c.drawString(350, y, f"ID No: {df.loc[i,'ID Number']}")
        c.drawString(50, y-15, f"KRA PIN: {df.loc[i,'KRA PIN']}")
        c.drawString(350, y-15, f"SHA No: {sha_no}")
        c.drawString(50, y-30, f"Phone: {df.loc[i,'Phone Number']}")
        y = height - 210
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, y, "EARNINGS"); c.drawString(250, y, "AMOUNT (KSH)")
        c.drawString(350, y, "DEDUCTIONS"); c.drawString(500, y, "AMOUNT (KSH)")
        c.line(50, y-3, width-50, y-3)
        c.setFont("Helvetica", 10)
        c.drawString(50, y-20, "Gross Salary"); c.drawRightString(320, y-20, f"{gross:,.2f}")
        c.drawString(350, y-20, "NSSF"); c.drawRightString(580, y-20, f"{nssf:,.2f}")
        c.drawString(350, y-35, "SHIF"); c.drawRightString(580, y-35, f"{shif:,.2f}")
        c.drawString(350, y-50, "Housing Levy"); c.drawRightString(580, y-50, f"{housing:,.2f}")
        c.drawString(350, y-65, "PAYE"); c.drawRightString(580, y-65, f"{paye:,.2f}")
        c.drawString(350, y-80, "Pension"); c.drawRightString(580, y-80, f"{pension:,.2f}")
        c.drawString(350, y-95, "HELB"); c.drawRightString(580, y-95, f"{helb:,.2f}")
        c.line(50, y-115, width-50, y-115)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y-135, "NET PAY"); c.drawRightString(320, y-135, f"{net_pay:,.2f}")
        c.setFont("Helvetica", 8)
        c.drawCentredString(width/2, 40, "This is a computer generated document")
        c.drawCentredString(width/2, 25, f"Generated on {datetime.now().strftime('%d/%m/%Y')}")
        c.save()
    print(f"✅ Generated {len(df)} PDFs in Payslips/")

def save_all(df, summary, phones_tab, cleartax):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name='Payroll', index=False)
        summary.to_excel(writer, sheet_name='KRA Summary', index=False)
        phones_tab.to_excel(writer, sheet_name='Phone List', index=False)
        cleartax.to_excel(writer, sheet_name='ClearTax CSV', index=False)
    cleartax.to_csv('ClearTax_Upload_2026.csv', index=False)
    print(f"✅ Saved {EXCEL_FILE}")

def send_whatsapp_messages(df, phones_tab):
    print(f"\n📱 Starting WhatsApp Sender for: {PAYROLL_MONTH}")
    print("Make sure you are logged in to web.whatsapp.com")
    input("Press ENTER to start...")
    for i in range(len(phones_tab)):
        name = phones_tab.loc[i, 'Employee Name']
        phone = str(phones_tab.loc[i, 'Phone Number'])
        phone_formatted = phone.replace("0", "254", 1) if phone.startswith("0") else phone
        net_pay = float(phones_tab.loc[i, 'Net Pay'])
        gross = float(df.loc[i,'Gross Salary'])
        message = f"Hi {name},\n\nPlease find attached your payslip for {PAYROLL_MONTH}.\n\n*Gross Salary*: Ksh {gross:,.2f}\n*Net Pay*: Ksh {net_pay:,.2f}\n\nThank you,\n{COMPANY_NAME}"
        url_message = urllib.parse.quote(message)
        url = f"https://wa.me/{phone_formatted}?text={url_message}"
        webbrowser.open(url)
        print(f"Opened WhatsApp for {name}")
        input("Press ENTER for next employee...")

def main():
    global df
    if len(sys.argv) < 2:
        print("Usage: python3 payroll.py create | sync | whatsapp")
        sys.exit()
    command = sys.argv[1].lower()
    if command == "create": df = create_new_file()
    elif command == "sync" or command == "whatsapp": df = load_from_excel()
    else: sys.exit()
    if command == "create" or command == "sync":
        df = add_formulas(df)
        summary, phones_tab, cleartax = build_other_tabs(df)
        save_all(df, summary, phones_tab, cleartax)
        generate_pdfs(df)
        print("\n🎉 All Done!")
    if command == "whatsapp":
        _, phones_tab, _ = build_other_tabs(df)
        send_whatsapp_messages(df, phones_tab)

if __name__ == "__main__":
    main()
