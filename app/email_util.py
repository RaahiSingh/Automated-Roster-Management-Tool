import pandas as pd
import re
import ssl
import smtplib
import os
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

def process_excel_and_send_email(filepath):
    df_raw = pd.read_excel(filepath, header=None)
    df = pd.read_excel(filepath, skiprows=3, header=None)

    data = []

    for _, row in df.iterrows():
        signum = row[1]
        name = row[4]
        if pd.isna(signum) or pd.isna(name):
            continue

        oc_days = []
        for col in range(6, len(row)):
            cell_value = str(row[col]).strip().upper()
            if cell_value == "OC":
                date = df_raw.iloc[2, col]
                oc_days.append(str(date))

        if oc_days:
            title_cell = str(df_raw.iloc[0].values)
            match = re.search(r"\b([A-Za-z]+)\s+(\d{4})", title_cell)
            if match:
                month, year = match.groups()
                month = month.title()
            else:
                month = "Invalid"

            oc_days_with_month = [f"{day}-{month}" for day in oc_days]
            data.append({
                "Name": name,
                "Signum": signum,
                "OnCall Support Days": f"{len(oc_days)} ({', '.join(oc_days_with_month)})",
                "Amount": round(len(oc_days) * 714.29, 2)
            })

    if not data:
        return False

    summary = pd.DataFrame(data)

    html_content = f"""
    <p>Hi Swati,</p>
    <p>Please find below the On Call summary:</p>
    <table style="font-family: 'Times New Roman', Times, serif; border-collapse: collapse; width: 80%; text-align: center;">
        <thead>
            <tr style="background-color: #f2a365; color: white;">
                <th style="border: 1px solid #999; padding: 8px;">Name of Team member</th>
                <th style="border: 1px solid #999; padding: 8px;">Signum</th>
                <th style="border: 1px solid #999; padding: 8px;">On Call support days</th>
                <th style="border: 1px solid #999; padding: 8px;">Amount</th>
            </tr>
        </thead>
        <tbody>
    """
    for _, row in summary.iterrows():
        html_content += f"""
            <tr style="background-color: white;">
                <td style="border: 1px solid #999; padding: 8px;">{row['Name']}</td>
                <td style="border: 1px solid #999; padding: 8px;">{row['Signum']}</td>
                <td style="border: 1px solid #999; padding: 8px;">{row['OnCall Support Days']}</td>
                <td style="border: 1px solid #999; padding: 8px;">{row['Amount']:.2f}</td>
            </tr>
        """
    html_content += "</tbody></table>"

    title_cell = str(df_raw.iloc[0].values)
    match = re.search(r"\b([A-Za-z]+)\s+(\d{4})", title_cell)
    subject = f"On Call Data – {match.group(1).title()}-{match.group(2)}" if match else "On Call Data – Unknown"

    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = EMAIL_RECEIVER
    msg["Subject"] = subject
    msg.attach(MIMEText(html_content, 'html'))

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"Email error: {e}")
        return False
