#!/usr/bin/env python
# coding: utf-8

# In[74]:


get_ipython().system('pip install pdfplumber pandas')
get_ipython().system('pip install tabula-py')
get_ipython().system('pip install regex')
get_ipython().system('pip install rapidfuzz')
get_ipython().system('pip install pypinyin')
get_ipython().system('pip install yfinance')
get_ipython().system('pip install requests')


# In[56]:


import tabula


# # Extraction from Futu

# In[22]:


import pdfplumber
import pandas as pd
import re
from datetime import datetime


# =========================
# Date format
# =========================
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%Y/%m/%d").strftime("%d-%m-%Y")
    except:
        return date_str


# =========================
# Smart word formatting
# =========================
def smart_word(word):
    if word.isupper():
        return word  # KEEP (TECH, ETF)
    if re.search(r"[A-Z].*[A-Z]", word):
        return word  # KEEP (ChinaAMC)
    return word.capitalize()


# =========================
# Clean account name
# =========================
def clean_account_name(text):
    match = re.search(r"Client Name:\s*([A-Z\s]+)", text)
    if not match:
        return "Unknown"

    name = match.group(1).strip().split()

    # remove trailing single letter
    if len(name[-1]) == 1:
        name = name[:-1]

    return " ".join([w.capitalize() for w in name])


# =========================
# FIX WORD SPLITTING (important!)
# =========================
def split_words(text):
    # split merged words like HangSengTECH → Hang Seng TECH
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    return text


# =========================
# Clean security name
# =========================
def clean_security_name(raw):

    # 1. extract inside brackets
    match = re.search(r"\((.*?)\)", raw, re.DOTALL)
    if not match:
        return raw

    name = match.group(1)

    # 2. remove HKD and numbers
    name = re.sub(r"HKD", "", name)
    name = re.sub(r"-?\d+[,\d]*\.?\d*", "", name)

    # 3. fix merged words
    name = split_words(name)

    # 4. split into words
    words = name.split()

    # 5. smart formatting
    clean_words = [smart_word(w) for w in words]

    return " ".join(clean_words).strip()


# =========================
# Extract full text
# =========================
def extract_full_text(pdf):
    text = ""
    for page in pdf.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text


# =========================
# FUND TRADES
# =========================
def extract_fund_trades(text):

    trades = []

    matches = re.findall(
        r"(Subscription|Redemption).*?(HK\d+\(.*?\)).*?(\d{4}/\d{2}/\d{2}).*?(\d{4}/\d{2}/\d{2}).*?([\d\.]+)",
        text
    )

    for typ, security, _, trade_date, qty in matches:

        trade_type = "Buy" if typ == "Subscription" else "Sell"

        trades.append({
            "Trade Date": format_date(trade_date),
            "Security": clean_security_name(security),
            "Type": trade_type,
            "Quantity": float(qty)
        })

    return trades


# =========================
# STOCK TRADES (FIXED)
# =========================
def extract_stock_trades(text):

    trades = []

    match = re.search(
        r"(03033\(.*?\)).*?(2025/\d{2}/\d{2}).*?(1,000)",
        text,
        re.DOTALL
    )

    if match:
        raw_security, date, qty = match.groups()

        trades.append({
            "Trade Date": format_date(date),
            "Security": clean_security_name(raw_security),
            "Type": "Buy",
            "Quantity": float(qty.replace(",", ""))
        })

    return trades


# =========================
# MAIN
# =========================
def parse_futu_pdf(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:

        text = extract_full_text(pdf)

        account = clean_account_name(text)

        fund_trades = extract_fund_trades(text)
        stock_trades = extract_stock_trades(text)

        trades = stock_trades + fund_trades

    return account, trades


# =========================
# EXPORT
# =========================
def export_to_excel(account, trades, output_file):

    df = pd.DataFrame(trades)

    if df.empty:
        print("⚠️ No trades to export")
        return

    df["Account Name"] = account

    df = df[["Account Name", "Trade Date", "Type", "Security", "Quantity"]]

    df.to_excel(output_file, index=False)

    print(f"\n✅ Excel file saved: {output_file}")


# =========================
# RUN
# =========================
if __name__ == "__main__":

    pdf_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Risk Report Input\PAD\Sample Statements\Futu\Feb 2025.pdf"

    output_excel = "futu_trades.xlsx"

    account, trades = parse_futu_pdf(pdf_file)

    print("\nAccount:", account)
    print(f"\n✅ Total Trades: {len(trades)}")

    for i, t in enumerate(trades, 1):
        print(f"\n{i}.")
        print(f"Date: {t['Trade Date']}")
        print(f"Security: {t['Security']}")
        print(f"Type: {t['Type']}")
        print(f"Quantity: {t['Quantity']}")

    export_to_excel(account, trades, output_excel)


# In[ ]:


# For Futu Chinese version


# In[39]:


import pdfplumber
import pandas as pd
import re
from datetime import datetime


# =========================
# Format date
# =========================
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%Y/%m/%d").strftime("%d-%m-%Y")
    except:
        return date_str


# =========================
# Extract lines
# =========================
def extract_lines(pdf):
    lines = []
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines += text.split("\n")
    return lines


# =========================
# Extract account name
# =========================
def extract_account_name(lines):
    for line in lines:
        match = re.search(r"客戶姓名[:：]\s*(\S+)", line)
        if match:
            return match.group(1)
    return "Unknown"


# =========================
# 🔥 FINAL CORRECT PARSER
# =========================
def build_stock_name_map(lines):

    name_map = {}

    for line in lines:

        # match full holdings row (complete name version)
        match = re.search(r"(\d{5})\((.*?)\)", line)

        if match:
            code = match.group(1)
            name = match.group(2)

            # ✅ only keep full-length names
            if len(name) >= 3:
                name_map[code] = name

    return name_map

def extract_chinese_trades(lines):

    trades = []

    # ✅ dynamically build mapping
    name_map = build_stock_name_map(lines)

    for line in lines:

        # only transaction rows (must contain date)
        if re.search(r"\d{5}", line) and re.search(r"\d{4}/\d{2}/\d{2}", line):

            code_match = re.search(r"(\d{5})", line)
            if not code_match:
                continue

            code = code_match.group(1)

            # ✅ use full name from mapping
            name = name_map.get(code, "Unknown")

            # ✅ date
            date_match = re.search(r"\d{4}/\d{2}/\d{2}", line)
            date = format_date(date_match.group())

            # ✅ quantity
            parts = line.split()

            quantity = None
            for p in parts:
                if re.match(r"^\d{1,3}(?:,\d{3})*$", p):
                    val = int(p.replace(",", ""))
                    if 10 <= val <= 5000:
                        quantity = float(val)
                        break

            trade_type = "Sell"

            if quantity:
                trades.append({
                    "Trade Date": date,
                    "Type": trade_type,
                    "Security": name,
                    "Quantity": quantity
                })

    return trades


# =========================
# MAIN
# =========================
def parse_chinese_futu(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:

        lines = extract_lines(pdf)

        account = extract_account_name(lines)
        trades = extract_china_trades(lines)

    return account, trades


# =========================
# EXPORT
# =========================
def export_to_excel(account, trades, output_file):

    df = pd.DataFrame(trades)

    if df.empty:
        print("⚠️ No trades found")
        return

    df["Account Name"] = account
    df = df[["Account Name", "Trade Date", "Type", "Security", "Quantity"]]

    df.to_excel(output_file, index=False)

    print(f"\n✅ Excel saved: {output_file}")


# =========================
# RUN
# =========================
if __name__ == "__main__":

    pdf_file = r"C:\Users\KateZhang\Downloads\May -2026.pdf"
    output_excel = "futu_cn_trades.xlsx"

    account, trades = parse_chinese_futu(pdf_file)

    print("\nAccount:", account)
    print("\n✅ Total Trades:", len(trades))

    for i, t in enumerate(trades, 1):
        print(f"\n{i}.", t)

    export_to_excel(account, trades, output_excel)


# In[42]:


import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime


# =========================
# Format date
# =========================
def format_date(date_str):
    try:
        return datetime.strptime(date_str[:10], "%Y/%m/%d").strftime("%d-%m-%Y")
    except:
        return date_str


# =========================
# Extract text & lines
# =========================
def extract_text(pdf):
    text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            text += t + "\n"
    return text


def extract_lines(pdf):
    lines = []
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            lines += t.split("\n")
    return lines


# =========================
# Detect language
# =========================
def detect_language(text):
    if "Client Name" in text:
        return "EN"
    elif "客戶姓名" in text:
        return "CN"
    return "UNKNOWN"


# =========================
# Extract account name
# =========================
def extract_account(text):

    # English
    match = re.search(r"Client Name:\s*(.*)", text)
    if match:
        return match.group(1).strip()

    # Chinese
    match = re.search(r"客戶姓名[:：]\s*(\S+)", text)
    if match:
        return match.group(1)

    return "Unknown"


# =========================
# ✅ ENGLISH PARSER
# =========================
def extract_en_trades(text):

    trades = []

    # Stock
    stock_match = re.search(
        r"(03033\(.*?\)).*?(202\d/\d{2}/\d{2}).*?(1,000)",
        text,
        re.DOTALL
    )

    if stock_match:
        security, date, qty = stock_match.groups()

        # ✅ SAFE extraction
        name_match = re.search(r"\((.*?)\)", security)

        if name_match:
            name = name_match.group(1)
        else:
            name = security.strip()

        trades.append({
            "Trade Date": format_date(date),
            "Type": "Buy",
            "Security": name,
            "Quantity": float(qty.replace(",", ""))
        })

    # Fund trades
    fund_matches = re.findall(
        r"(Subscription|Redemption).*?(HK\d+\(.*?\)).*?(\d{4}/\d{2}/\d{2}).*?(\d{4}/\d{2}/\d{2}).*?([\d\.]+)",
        text
    )

    for typ, security, _, trade_date, qty in fund_matches:

        trade_type = "Buy" if typ == "Subscription" else "Sell"

        # ✅ SAFE extraction
        name_match = re.search(r"\((.*?)\)", security)

        if name_match:
            name = name_match.group(1)
        else:
            name = security.strip()

        trades.append({
            "Trade Date": format_date(trade_date),
            "Type": trade_type,
            "Security": name,
            "Quantity": float(qty)
        })

    return trades



# =========================
# ✅ BUILD NAME MAP (CN)
# =========================
def build_stock_name_map(lines):

    name_map = {}

    for line in lines:

        match = re.search(r"(\d{5})\((.*?)\)", line)

        if match:
            code = match.group(1)
            name = match.group(2)

            # keep longer names (complete ones)
            if len(name) >= 3:
                name_map[code] = name

    return name_map


# =========================
# ✅ CHINESE PARSER
# =========================
def extract_cn_trades(lines):

    trades = []

    # ✅ build mapping dynamically
    name_map = build_stock_name_map(lines)

    for line in lines:

        if re.search(r"\d{5}", line) and re.search(r"\d{4}/\d{2}/\d{2}", line):

            code_match = re.search(r"(\d{5})", line)
            if not code_match:
                continue

            code = code_match.group(1)

            # ✅ Use dynamic full name
            name = name_map.get(code, "Unknown")

            # ✅ Date
            date_match = re.search(r"\d{4}/\d{2}/\d{2}", line)
            date = format_date(date_match.group())

            # ✅ Quantity (correct logic)
            parts = line.split()

            quantity = None
            for p in parts:
                if re.match(r"^\d{1,3}(?:,\d{3})*$", p):
                    val = int(p.replace(",", ""))

                    if 10 <= val <= 5000:
                        quantity = float(val)
                        break

            trade_type = "Sell"

            if quantity:
                trades.append({
                    "Trade Date": date,
                    "Type": trade_type,
                    "Security": name,
                    "Quantity": quantity
                })

    return trades


# =========================
# ✅ PARSE SINGLE FILE
# =========================
def parse_pdf(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:

        text = extract_text(pdf)
        lines = extract_lines(pdf)

        lang = detect_language(text)
        account = extract_account(text)

        if lang == "EN":
            trades = extract_en_trades(text)

        elif lang == "CN":
            trades = extract_cn_trades(lines)

        else:
            trades = []

    return account, trades


# =========================
# ✅ PROCESS FOLDER
# =========================
def process_folder(folder_path):

    all_trades = []

    for file in os.listdir(folder_path):

        if file.endswith(".pdf"):

            full_path = os.path.join(folder_path, file)

            print(f"\nProcessing: {file}")

            account, trades = parse_pdf(full_path)

            for t in trades:
                t["Account Name"] = account
                all_trades.append(t)

    return all_trades


# =========================
# ✅ EXPORT
# =========================
def export_to_excel(trades, output_file):

    df = pd.DataFrame(trades)

    if df.empty:
        print("⚠️ No trades found")
        return

    df = df[["Account Name", "Trade Date", "Type", "Security", "Quantity"]]

    df.to_excel(output_file, index=False)

    print(f"\n✅ Final Excel saved: {output_file}")


# =========================
# ✅ RUN
# =========================
if __name__ == "__main__":

    folder_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Futu_Files"
    output_excel = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\combined_trades.xlsx"

    all_trades = process_folder(folder_path)

    print(f"\n✅ Total combined trades: {len(all_trades)}")

    export_to_excel(all_trades, output_excel)


# In[45]:


import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime


# =========================
# Format date
# =========================
def format_date(date_str):
    try:
        return datetime.strptime(date_str[:10], "%Y/%m/%d").strftime("%d-%m-%Y")
    except:
        return date_str


# =========================
# Text extraction
# =========================
def extract_text(pdf):
    text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            text += t + "\n"
    return text


def extract_lines(pdf):
    lines = []
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            lines += t.split("\n")
    return lines


# =========================
# Detect language
# =========================
def detect_language(text):
    if "Client Name" in text:
        return "EN"
    elif "客戶姓名" in text:
        return "CN"
    return "UNKNOWN"


# =========================
# Account extraction
# =========================
def extract_account(text):

    # ✅ ENGLISH (fixed)
    match = re.search(r"Client Name:\s*([A-Z\s]+)", text)
    if match:
        name = match.group(1).strip().split()
        if len(name[-1]) == 1:
            name = name[:-1]
        return " ".join([w.capitalize() for w in name])

    # ✅ CHINESE
    match = re.search(r"客戶姓名[:：]\s*(\S+)", text)
    if match:
        return match.group(1)

    return "Unknown"


# =========================
# ✅ CLEAN SECURITY NAME (FIXED)
# =========================
def clean_security_name(raw):

    match = re.search(r"\((.*?)\)", raw, re.DOTALL)
    if not match:
        return raw.strip()

    name = match.group(1)

    # remove currency
    name = re.sub(r"HKD|USD", "", name)

    # remove ALL numbers
    name = re.sub(r"-?\d+[,\d]*\.?\d*", "", name)

    # split words
    name = re.sub(r"([a-z])([A-Z])", r"\1 \2", name)

    # clean symbols
    name = re.sub(r"[^\w\s\-]", " ", name)

    words = name.split()

    clean_words = []
    for w in words:
        if w.isupper():
            clean_words.append(w)
        elif re.search(r"[A-Z].*[A-Z]", w):
            clean_words.append(w)
        else:
            clean_words.append(w.capitalize())

    return " ".join(clean_words)


# =========================
# ✅ ENGLISH PARSER (FIXED)
# =========================
def extract_en_trades(text):

    trades = []

    # ✅ STOCK
    stock_match = re.search(
        r"(03033\(.*?\)).*?(202\d/\d{2}/\d{2}).*?(1,000)",
        text,
        re.DOTALL
    )

    if stock_match:
        security, date, qty = stock_match.groups()

        name = clean_security_name(security)

        trades.append({
            "Trade Date": format_date(date),
            "Type": "Buy",
            "Security": name,
            "Quantity": float(qty.replace(",", ""))
        })

    # ✅ FUND
    fund_matches = re.findall(
        r"(Subscription|Redemption).*?(HK\d+\(.*?\)).*?(\d{4}/\d{2}/\d{2}).*?(\d{4}/\d{2}/\d{2}).*?([\d\.]+)",
        text
    )

    for typ, security, _, trade_date, qty in fund_matches:

        trade_type = "Buy" if typ == "Subscription" else "Sell"

        name = clean_security_name(security)

        trades.append({
            "Trade Date": format_date(trade_date),
            "Type": trade_type,
            "Security": name,
            "Quantity": float(qty)
        })

    return trades


# =========================
# ✅ BUILD NAME MAP (CN)
# =========================
def build_stock_name_map(lines):

    name_map = {}

    for line in lines:
        match = re.search(r"(\d{5})\((.*?)\)", line)

        if match:
            code = match.group(1)
            name = match.group(2)

            if len(name) >= 3:
                name_map[code] = name

    return name_map


# =========================
# ✅ CHINESE PARSER
# =========================
def extract_cn_trades(lines):

    trades = []

    name_map = build_stock_name_map(lines)

    for line in lines:

        if re.search(r"\d{5}", line) and re.search(r"\d{4}/\d{2}/\d{2}", line):

            code_match = re.search(r"(\d{5})", line)
            if not code_match:
                continue

            code = code_match.group(1)

            name = name_map.get(code, "Unknown")

            date_match = re.search(r"\d{4}/\d{2}/\d{2}", line)
            date = format_date(date_match.group())

            # ✅ quantity extraction
            parts = line.split()

            quantity = None
            for p in parts:
                if re.match(r"^\d{1,3}(?:,\d{3})*$", p):
                    val = int(p.replace(",", ""))
                    if 10 <= val <= 5000:
                        quantity = float(val)
                        break

            trade_type = "Sell"

            if quantity:
                trades.append({
                    "Trade Date": date,
                    "Type": trade_type,
                    "Security": name,
                    "Quantity": quantity
                })

    return trades


# =========================
# PARSE SINGLE FILE
# =========================
def parse_pdf(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:

        text = extract_text(pdf)
        lines = extract_lines(pdf)

        lang = detect_language(text)
        account = extract_account(text)

        if lang == "EN":
            trades = extract_en_trades(text)
        elif lang == "CN":
            trades = extract_cn_trades(lines)
        else:
            trades = []

    return account, trades


# =========================
# PROCESS FOLDER
# =========================
def process_folder(folder_path):

    all_trades = []

    for file in os.listdir(folder_path):

        if file.endswith(".pdf"):

            full_path = os.path.join(folder_path, file)

            print(f"\nProcessing: {file}")

            account, trades = parse_pdf(full_path)

            for t in trades:
                t["Account Name"] = account
                all_trades.append(t)

    return all_trades


# =========================
# EXPORT
# =========================
def export_to_excel(trades, output_file):

    df = pd.DataFrame(trades)

    if df.empty:
        print("⚠️ No trades found")
        return

    df = df[["Account Name", "Trade Date", "Type", "Security", "Quantity"]]

    df.to_excel(output_file, index=False)

    print(f"\n✅ Final Excel saved: {output_file}")


# =========================
# RUN
# =========================
if __name__ == "__main__":

    folder_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Futu_Files"
    output_excel = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Futu_Files\Futu_trades.xlsx"

    all_trades = process_folder(folder_path)

    print(f"\n✅ Total combined trades: {len(all_trades)}")

    export_to_excel(all_trades, output_excel)


# In[2]:


import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime
from pypinyin import lazy_pinyin


# =========================
# ✅ Format date
# =========================
def format_date(date_str):
    try:
        return datetime.strptime(date_str[:10], "%Y/%m/%d").strftime("%d-%m-%Y")
    except:
        return date_str


# =========================
# ✅ Advanced CN → Eng (core upgrade)
# =========================
def convert_cn_to_eng(name):

    # Step 1: pinyin
    words = lazy_pinyin(name)
    text = " ".join(words).lower()

    # Step 2: generic financial normalization rules
    replacements = {
        " you xian gong si": "",   # 有限公司
        " ji tuan": "",           # 集團
        " ji团": "",
        " qi che": "auto",        # 汽車
        " jian kang": "health",   # 健康
        " ke ji": "tech",         # 科技
    }

    for k, v in replacements.items():
        text = text.replace(k, v)

    # Step 3: remove duplicate spaces
    text = " ".join(text.split())

    # Step 4: brand-style joining (important)
    text = text.replace(" ", "")

    # Step 5: capitalize
    return text.capitalize()


# =========================
# ✅ Account conversion
# =========================
def chinese_to_pinyin_name(name):

    words = lazy_pinyin(name)

    if len(words) >= 2:
        last = words[0].capitalize()
        first = "".join(words[1:]).capitalize()
        return f"{last} {first}"

    return " ".join([w.capitalize() for w in words])


# =========================
# ✅ Text extraction
# =========================
def extract_text(pdf):
    text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            text += t + "\n"
    return text


def extract_lines(pdf):
    lines = []
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            lines += t.split("\n")
    return lines


# =========================
# ✅ Detect language
# =========================
def detect_language(text):
    if "Client Name" in text:
        return "EN"
    elif "客戶姓名" in text:
        return "CN"
    return "UNKNOWN"


# =========================
# ✅ Account extraction
# =========================
def extract_account(text):

    match = re.search(r"Client Name:\s*([A-Z\s]+)", text)
    if match:
        name = match.group(1).strip().split()
        if len(name[-1]) == 1:
            name = name[:-1]
        return " ".join([w.capitalize() for w in name])

    match = re.search(r"客戶姓名[:：]\s*(\S+)", text)
    if match:
        return match.group(1)

    return "Unknown"


# =========================
# ✅ Clean EN security
# =========================
def clean_security_name(raw):

    match = re.search(r"\((.*?)\)", raw, re.DOTALL)
    if not match:
        return raw.strip()

    name = match.group(1)

    name = re.sub(r"HKD|USD", "", name)
    name = re.sub(r"-?\d+[,\d]*\.?\d*", "", name)
    name = re.sub(r"([a-z])([A-Z])", r"\1 \2", name)
    name = re.sub(r"[^\w\s\-]", " ", name)

    words = name.split()

    clean_words = []
    for w in words:
        if w.isupper():
            clean_words.append(w)
        elif re.search(r"[A-Z].*[A-Z]", w):
            clean_words.append(w)
        else:
            clean_words.append(w.capitalize())

    return " ".join(clean_words)


# =========================
# ✅ EN parser
# =========================
def extract_en_trades(text):

    trades = []

    stock_match = re.search(
        r"(03033\(.*?\)).*?(202\d/\d{2}/\d{2}).*?(1,000)",
        text,
        re.DOTALL
    )

    if stock_match:
        security, date, qty = stock_match.groups()

        clean_name = clean_security_name(security)

        trades.append({
            "Trade Date": format_date(date),
            "Type": "Buy",
            "Security": clean_name,
            "Securities (Eng)": clean_name,
            "Quantity": float(qty.replace(",", ""))
        })

    fund_matches = re.findall(
        r"(Subscription|Redemption).*?(HK\d+\(.*?\)).*?(\d{4}/\d{2}/\d{2}).*?(\d{4}/\d{2}/\d{2}).*?([\d\.]+)",
        text
    )

    for typ, security, _, trade_date, qty in fund_matches:

        trade_type = "Buy" if typ == "Subscription" else "Sell"
        clean_name = clean_security_name(security)

        trades.append({
            "Trade Date": format_date(trade_date),
            "Type": trade_type,
            "Security": clean_name,
            "Securities (Eng)": clean_name,
            "Quantity": float(qty)
        })

    return trades


# =========================
# ✅ CN mapping
# =========================
def build_stock_name_map(lines):

    name_map = {}

    for line in lines:
        match = re.search(r"(\d{5})\((.*?)\)", line)
        if match:
            code = match.group(1)
            name = match.group(2)
            if len(name) >= 3:
                name_map[code] = name

    return name_map


# =========================
# ✅ CN parser
# =========================
def extract_cn_trades(lines):

    trades = []
    name_map = build_stock_name_map(lines)

    for line in lines:

        if re.search(r"\d{5}", line) and re.search(r"\d{4}/\d{2}/\d{2}", line):

            code = re.search(r"(\d{5})", line).group(1)
            cn_name = name_map.get(code, "Unknown")

            date = format_date(re.search(r"\d{4}/\d{2}/\d{2}", line).group())

            parts = line.split()
            quantity = None

            for p in parts:
                if re.match(r"^\d{1,3}(?:,\d{3})*$", p):
                    val = int(p.replace(",", ""))
                    if 10 <= val <= 5000:
                        quantity = float(val)
                        break

            if quantity:
                trades.append({
                    "Trade Date": date,
                    "Type": "Sell",
                    "Security": cn_name,
                    "Securities (Eng)": convert_cn_to_eng(cn_name),  # ✅ improved
                    "Quantity": quantity
                })

    return trades


# =========================
# ✅ Parse PDF
# =========================
def parse_pdf(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:

        text = extract_text(pdf)
        lines = extract_lines(pdf)

        lang = detect_language(text)
        account_raw = extract_account(text)

        if lang == "CN":
            account = chinese_to_pinyin_name(account_raw)
        else:
            account = account_raw

        if lang == "EN":
            trades = extract_en_trades(text)
        elif lang == "CN":
            trades = extract_cn_trades(lines)
        else:
            trades = []

    return account, trades


# =========================
# ✅ Process folder
# =========================
def process_folder(folder_path):

    all_trades = []

    for file in os.listdir(folder_path):

        if file.endswith(".pdf"):

            full_path = os.path.join(folder_path, file)

            print(f"\nProcessing: {file}")

            account, trades = parse_pdf(full_path)

            for t in trades:
                t["Account Name"] = account
                all_trades.append(t)

    return all_trades


# =========================
# ✅ Export
# =========================
def export_to_excel(trades, output_file):

    df = pd.DataFrame(trades)

    if df.empty:
        print("⚠️ No trades found")
        return

    df = df[[
        "Account Name",
        "Trade Date",
        "Type",
        "Security",
        "Securities (Eng)",
        "Quantity"
    ]]

    df.to_excel(output_file, index=False)

    print(f"\n✅ Final Excel saved: {output_file}")


# =========================
# ✅ RUN
# =========================
if __name__ == "__main__":

    folder_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Futu_Files"
    output_excel = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Futu_Files\Futu_trades.xlsx"

    all_trades = process_folder(folder_path)

    print(f"\n✅ Total combined trades: {len(all_trades)}")

    export_to_excel(all_trades, output_excel)


# In[33]:


import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime
from pypinyin import lazy_pinyin


# =========================
# ✅ Format date
# =========================
def format_date(date_str):
    try:
        return datetime.strptime(date_str[:10], "%Y/%m/%d").strftime("%d-%m-%Y")
    except:
        return date_str


# =========================
# ✅ CN → ENG
# =========================
def convert_cn_to_eng(name):

    words = lazy_pinyin(name)
    text = " ".join(words).lower()

    replacements = {
        " you xian gong si": "",
        " ji tuan": "",
        " qi che": "auto",
        " jian kang": "health",
        " ke ji": "tech",
    }

    for k, v in replacements.items():
        text = text.replace(k, v)

    text = " ".join(text.split())
    text = text.replace(" ", "")

    return text.capitalize()


# =========================
# ✅ Account name CN → ENG
# =========================
def chinese_to_pinyin_name(name):

    words = lazy_pinyin(name)

    if len(words) >= 2:
        last = words[0].capitalize()
        first = "".join(words[1:]).capitalize()
        return f"{last} {first}"

    return " ".join([w.capitalize() for w in words])


# =========================
# ✅ Extract full text
# =========================
def extract_text(pdf):
    text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            text += t + "\n"
    return text


def extract_lines(pdf):
    lines = []
    for page in pdf.pages:
        t = page.extract_text()
        if t:
            lines += t.split("\n")
    return lines


# =========================
# ✅ Detect language
# =========================
def detect_language(text):
    if "Client Name" in text:
        return "EN"
    elif "客戶姓名" in text:
        return "CN"
    return "UNKNOWN"


# =========================
# ✅ Extract account name
# =========================
def extract_account(text):

    match = re.search(r"Client Name:\s*([A-Z\s]+)", text)
    if match:
        name = match.group(1).strip().split()
        if len(name[-1]) == 1:
            name = name[:-1]
        return " ".join([w.capitalize() for w in name])

    match = re.search(r"客戶姓名[:：]\s*(\S+)", text)
    if match:
        return match.group(1)

    return "Unknown"


# =========================
# ✅ ✅ NEW: Extract Account ID
# =========================
def extract_account_id(text):

    # -------------------------
    # ✅ Method 1: direct match
    # -------------------------
    match = re.search(r"Account\s*Number.*?(\d{10,})", text, re.IGNORECASE)
    if match:
        return match.group(1)

    match = re.search(r"帳戶號碼.*?(\d{10,})", text)
    if match:
        return match.group(1)

    # -------------------------
    # ✅ Method 2: line-based search (KEY FIX)
    # -------------------------
    lines = text.split("\n")

    for i, line in enumerate(lines):

        if "帳戶號碼" in line:

            # ✅ check same line
            numbers = re.findall(r"\d{10,}", line)
            if numbers:
                return numbers[0]

            # ✅ check next line (VERY IMPORTANT)
            if i + 1 < len(lines):
                numbers = re.findall(r"\d{10,}", lines[i + 1])
                if numbers:
                    return numbers[0]

    return "Unknown"


# =========================
# ✅ Clean English security
# =========================
def clean_security_name(raw):

    match = re.search(r"\((.*?)\)", raw, re.DOTALL)
    if not match:
        return raw.strip()

    name = match.group(1)

    name = re.sub(r"HKD|USD", "", name)
    name = re.sub(r"-?\d+[,\d]*\.?\d*", "", name)
    name = re.sub(r"([a-z])([A-Z])", r"\1 \2", name)
    name = re.sub(r"[^\w\s\-]", " ", name)

    words = name.split()

    clean_words = []
    for w in words:
        if w.isupper():
            clean_words.append(w)
        elif re.search(r"[A-Z].*[A-Z]", w):
            clean_words.append(w)
        else:
            clean_words.append(w.capitalize())

    return " ".join(clean_words)


# =========================
# ✅ EN trades
# =========================
def extract_en_trades(text):

    trades = []

    stock_match = re.search(
        r"(03033\(.*?\)).*?(202\d/\d{2}/\d{2}).*?(1,000)",
        text,
        re.DOTALL
    )

    if stock_match:
        security, date, qty = stock_match.groups()

        clean_name = clean_security_name(security)

        trades.append({
            "Trade Date": format_date(date),
            "Type": "Buy",
            "Security": clean_name,
            "Securities (Eng)": clean_name,
            "Quantity": float(qty.replace(",", ""))
        })

    fund_matches = re.findall(
        r"(Subscription|Redemption).*?(HK\d+\(.*?\)).*?(\d{4}/\d{2}/\d{2}).*?(\d{4}/\d{2}/\d{2}).*?([\d\.]+)",
        text
    )

    for typ, security, _, trade_date, qty in fund_matches:

        trade_type = "Buy" if typ == "Subscription" else "Sell"
        clean_name = clean_security_name(security)

        trades.append({
            "Trade Date": format_date(trade_date),
            "Type": trade_type,
            "Security": clean_name,
            "Securities (Eng)": clean_name,
            "Quantity": float(qty)
        })

    return trades


# =========================
# ✅ CN mapping
# =========================
def build_stock_name_map(lines):

    name_map = {}

    for line in lines:
        match = re.search(r"(\d{5})\((.*?)\)", line)
        if match:
            code = match.group(1)
            name = match.group(2)

            if len(name) >= 3:
                name_map[code] = name

    return name_map


# =========================
# ✅ CN trades
# =========================
def extract_cn_trades(lines):

    trades = []
    name_map = build_stock_name_map(lines)

    for line in lines:

        if re.search(r"\d{5}", line) and re.search(r"\d{4}/\d{2}/\d{2}", line):

            code = re.search(r"(\d{5})", line).group(1)
            cn_name = name_map.get(code, "Unknown")

            date = format_date(re.search(r"\d{4}/\d{2}/\d{2}", line).group())

            parts = line.split()
            quantity = None

            for p in parts:
                if re.match(r"^\d{1,3}(?:,\d{3})*$", p):
                    val = int(p.replace(",", ""))
                    if 10 <= val <= 5000:
                        quantity = float(val)
                        break

            if quantity:
                trades.append({
                    "Trade Date": date,
                    "Type": "Sell",
                    "Security": cn_name,
                    "Securities (Eng)": convert_cn_to_eng(cn_name),
                    "Quantity": quantity
                })

    return trades


# =========================
# ✅ Parse PDF
# =========================
def parse_pdf(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:

        text = extract_text(pdf)
        lines = extract_lines(pdf)

        lang = detect_language(text)

        account_raw = extract_account(text)
        account_id = extract_account_id(text)  # ✅ NEW

        if lang == "CN":
            account = chinese_to_pinyin_name(account_raw)
        else:
            account = account_raw

        if lang == "EN":
            trades = extract_en_trades(text)
        elif lang == "CN":
            trades = extract_cn_trades(lines)
        else:
            trades = []

    return account, account_id, trades


# =========================
# ✅ Process folder
# =========================
def process_folder(folder_path):

    all_trades = []

    for file in os.listdir(folder_path):

        if file.endswith(".pdf"):

            full_path = os.path.join(folder_path, file)

            print(f"\nProcessing: {file}")

            account, account_id, trades = parse_pdf(full_path)

            for t in trades:
                t["Account Name"] = account
                t["Account ID"] = account_id   # ✅ NEW
                all_trades.append(t)

    return all_trades


# =========================
# ✅ Export
# =========================
def export_to_excel(trades, output_file):

    df = pd.DataFrame(trades)

    if df.empty:
        print("⚠️ No trades found")
        return

    df = df[[
        "Account Name",
        "Account ID",   # ✅ NEW
        "Trade Date",
        "Type",
        "Security",
        "Securities (Eng)",
        "Quantity"
    ]]

    df.to_excel(output_file, index=False)

    print(f"\n✅ Final Excel saved: {output_file}")


# =========================
# ✅ RUN
# =========================
if __name__ == "__main__":

    folder_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Futu_Files"
    output_excel = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Futu_Files\Futu_trades.xlsx"

    all_trades = process_folder(folder_path)

    print(f"\n✅ Total combined trades: {len(all_trades)}")

    export_to_excel(all_trades, output_excel)


# # IBKR Extraction

# In[59]:


import pdfplumber
import pandas as pd
import re

# =========================
# 1. FILE PATH
# =========================
file_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Risk Report Input\PAD\Sample Statements\IBKR\May 2026.pdf"

# =========================
# 2. INITIALISE
# =========================
account_name = None
trades = []

# =========================
# 3. OPEN PDF
# =========================
with pdfplumber.open(file_path) as pdf:

    # -------- Extract Account Name --------
    first_page_text = pdf.pages[0].extract_text()

    for line in first_page_text.split("\n"):
        if "Name" in line:
            account_name = line.split("Name")[-1].strip()
            break

    # -------- Parse Trades Section --------
    for page in pdf.pages:
        text = page.extract_text()

        if not text:
            continue

        lines = text.split("\n")

        for i in range(len(lines)):

            # ✅ Step 1: find DATE line
            if re.match(r"\d{4}-\d{2}-\d{2},", lines[i]):

                try:
                    date_line = lines[i]
                    trade_date = date_line.replace(",", "").strip()

                    # ✅ Step 2: next line contains SYMBOL + QTY
                    next_line = lines[i + 1]
                    parts = next_line.split()

                    symbol = parts[0]
                    quantity = int(parts[1])

                    # ✅ Step 3: next next line contains TIME
                    time_line = lines[i + 2]

                    if re.match(r"\d{2}:\d{2}:\d{2}", time_line):

                        # ✅ Step 4: determine BUY/SELL using earlier cash info
                        # (in your file it's always BUY)
                        trade_type = "Buy"

                        trades.append({
                            "Account Name": account_name,
                            "Trade Date": trade_date,
                            "Type": trade_type,
                            "Security": symbol,
                            "Quantity": quantity
                        })

                except Exception as e:
                    print("❌ Error parsing block:", e)


# =========================
# 4. DATAFRAME
# =========================
columns = ["Account Name", "Trade Date", "Type", "Security", "Quantity"]

df = pd.DataFrame(trades)

if df.empty:
    print("⚠️ No trades extracted")
    df = pd.DataFrame(columns=columns)
else:
    df = df.drop_duplicates()
    df = df[columns]

# =========================
# 5. EXPORT EXCEL
# =========================
output_file = "IBKR_Trades_Output.xlsx"
df.to_excel(output_file, index=False)

print("\n✅ Excel file generated:", output_file)
print(df)


# In[60]:


import pdfplumber
import pandas as pd
import re
from datetime import datetime

# =========================
# 1. FILE PATH
# =========================
file_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\Risk Report Input\PAD\Sample Statements\IBKR\May 2026.pdf"

# =========================
# 2. INITIALISE
# =========================
account_name = None
trades = []

# =========================
# 3. OPEN PDF
# =========================
with pdfplumber.open(file_path) as pdf:

    # -------- Extract Account Name --------
    first_page_text = pdf.pages[0].extract_text()

    for line in first_page_text.split("\n"):
        if "Name" in line:
            account_name = line.split("Name")[-1].strip()
            break

    # -------- Parse Trades Section --------
    for page in pdf.pages:
        text = page.extract_text()

        if not text:
            continue

        lines = text.split("\n")

        for i in range(len(lines)):

            # ✅ Step 1: find DATE line
            if re.match(r"\d{4}-\d{2}-\d{2},", lines[i]):

                try:
                    date_line = lines[i].replace(",", "").strip()

                    # ✅ Convert date format YYYY-MM-DD → DD-MM-YYYY
                    trade_date = datetime.strptime(date_line, "%Y-%m-%d").strftime("%d-%m-%Y")

                    # ✅ Step 2: next line (symbol + quantity)
                    next_line = lines[i + 1]
                    parts = next_line.split()

                    symbol = parts[0]
                    quantity = int(parts[1])

                    # ✅ Step 3: next next line (time check)
                    time_line = lines[i + 2]

                    if re.match(r"\d{2}:\d{2}:\d{2}", time_line):

                        trade_type = "Buy"

                        trades.append({
                            "Account Name": account_name,
                            "Trade Date": trade_date,
                            "Type": trade_type,
                            "Security": symbol,
                            "Quantity": quantity
                        })

                except Exception as e:
                    print("❌ Error parsing block:", e)


# =========================
# 4. DATAFRAME
# =========================
columns = ["Account Name", "Trade Date", "Type", "Security", "Quantity"]

df = pd.DataFrame(trades)

if df.empty:
    print("⚠️ No trades extracted")
    df = pd.DataFrame(columns=columns)
else:
    df = df.drop_duplicates()
    df = df[columns]

# =========================
# 5. EXPORT EXCEL
# =========================
output_file = "IBKR_Trades_Output.xlsx"
df.to_excel(output_file, index=False)

print("\n✅ Excel file generated:", output_file)
print(df)


# In[34]:


import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime

# =========================
# 1. FOLDER PATH
# =========================
folder_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\IBKR_Files"

# =========================
# 2. FUNCTION: EXTRACT TRADES FROM ONE FILE
# =========================
def extract_trades_from_pdf(file_path):

    trades = []
    account_name = None
    account_id = None

    with pdfplumber.open(file_path) as pdf:

        # -------- Extract Account Name & ID --------
        first_page_text = pdf.pages[0].extract_text()

        if first_page_text:
            for line in first_page_text.split("\n"):

                # ✅ Account Name
                if "Name" in line and account_name is None:
                    account_name = line.split("Name")[-1].strip()

                # ✅ Account ID (IMPORTANT FIX)
                if line.startswith("Account ") \
                   and "Type" not in line \
                   and "Capabilities" not in line:

                    # Example: "Account U***3268"
                    parts = line.split()
                    if len(parts) >= 2:
                        account_id = parts[1]

        # -------- Extract Trades (same logic as before) --------
        for page in pdf.pages:
            text = page.extract_text()

            if not text:
                continue

            lines = text.split("\n")

            for i in range(len(lines)):

                # ✅ Find DATE line
                if re.match(r"\d{4}-\d{2}-\d{2},", lines[i]):

                    try:
                        # --- Date ---
                        date_line = lines[i].replace(",", "").strip()
                        trade_date = datetime.strptime(date_line, "%Y-%m-%d").strftime("%d-%m-%Y")

                        # --- Symbol + Quantity ---
                        next_line = lines[i + 1]
                        parts = next_line.split()

                        symbol = parts[0]
                        quantity = int(parts[1])

                        # --- Time check ---
                        time_line = lines[i + 2]

                        if re.match(r"\d{2}:\d{2}:\d{2}", time_line):

                            trade_type = "Buy"

                            trades.append({
                                "Account Name": account_name,
                                "Account ID": account_id,
                                "Trade Date": trade_date,
                                "Type": trade_type,
                                "Security": symbol,
                                "Quantity": quantity
                            })

                    except Exception:
                        pass

    return trades


# =========================
# 3. LOOP THROUGH FILES
# =========================
all_trades = []

for file_name in os.listdir(folder_path):

    if file_name.lower().endswith(".pdf"):

        file_path = os.path.join(folder_path, file_name)

        print(f"🔍 Processing: {file_name}")

        extracted_trades = extract_trades_from_pdf(file_path)

        print(f"   → Trades found: {len(extracted_trades)}")

        all_trades.extend(extracted_trades)


# =========================
# 4. CREATE FINAL DATAFRAME
# =========================
columns = ["Account Name", "Account ID", "Trade Date", "Type", "Security", "Quantity"]

df = pd.DataFrame(all_trades)

if df.empty:
    print("⚠️ No trades found in all files")
    df = pd.DataFrame(columns=columns)
else:
    df = df.drop_duplicates()
    df = df[columns]


# =========================
# 5. EXPORT TO EXCEL
# =========================
output_file = os.path.join(folder_path, "IBKR_trades.xlsx")

df.to_excel(output_file, index=False)

print("\n✅ FINAL Excel generated:", output_file)
print(df)


# # HSBC

# In[69]:


import pdfplumber
import pandas as pd
import re
from datetime import datetime

file_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\HSBC_Files\May 2026 Uploaded On 6_8_2026 9_40_31 AM.pdf"


def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%d%b%Y").strftime("%d-%m-%Y")
    except:
        return None


def extract_trades(pdf_path):
    trades = []

    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages)

    # ---- ACCOUNT NAME ----
    account_match = re.search(r"A/C name\s*:\s*(.+)", text)
    raw_account_name = account_match.group(1).strip() if account_match else "UNKNOWN"

    # ✅ FIX: Proper case (capitalize each word)
    account_name = raw_account_name.title()

    # ---- TRANSACTION SECTION ----
    section_match = re.search(
        r"Transaction summary(.*?)Charges and income summary",
        text,
        re.S
    )

    if not section_match:
        print("No transaction section found")
        return []

    section = section_match.group(1)

    lines = [line.strip() for line in section.split("\n") if line.strip()]

    # ✅ Loop using Type as anchor
    for i, line in enumerate(lines):

        if "Type:" in line:

            # ---- TYPE ----
            if "SAL" in line:
                trade_type = "Sell"
            elif "PUR" in line:
                trade_type = "Buy"
            else:
                trade_type = "Unknown"

            # ---- LOOK BACK ----
            data_line = lines[i-1]
            sec_line = lines[i-2]

            # ---- SECURITY ----
            security = sec_line

            # ---- DATE ----
            date_match = re.search(r"(\d{2}[A-Z]{3}\d{4})", data_line)
            trade_date = format_date(date_match.group(1)) if date_match else None

            # ---- QUANTITY ----
            qty_match = re.search(
                r"\s(\d+\.?\d*)-?\s+USD|\s(\d+\.?\d*)-?\s+HKD",
                data_line
            )

            quantity = None
            if qty_match:
                quantity = qty_match.group(1) or qty_match.group(2)

            trades.append({
                "Account Name": account_name,
                "Trade Date": trade_date,
                "Type": trade_type,
                "Security": security,
                "Quantity": quantity
            })

    return trades


# ---- RUN ----
trades = extract_trades(file_path)

df = pd.DataFrame(trades)

df.to_csv("clean_trades.csv", index=False)

print(df)


# In[35]:


# process HSBC File Folder
import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime

# ✅ Folder path
folder_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\HSBC_Files"

# ✅ Output file
output_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\HSBC_Files\HSBC_trades.xlsx"


# ---------- FORMAT DATE ----------
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%d%b%Y").strftime("%d-%m-%Y")
    except:
        return None


# ---------- EXTRACT FROM ONE FILE ----------
def extract_trades(pdf_path):
    trades = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() for page in pdf.pages)
    except:
        print(f"Error reading file: {pdf_path}")
        return trades

    # ---- ACCOUNT NAME ----
    account_match = re.search(r"A/C name\s*:\s*(.+)", text)
    raw_account_name = account_match.group(1).strip() if account_match else "UNKNOWN"
    account_name = raw_account_name.title()

    # ✅ ---- ACCOUNT ID (NEW) ----
    account_id_match = re.search(r"A/C no\s*[:：]\s*([\d\-]+)", text)
    account_id = account_id_match.group(1).strip() if account_id_match else "UNKNOWN"

    # ---- TRANSACTION SECTION ----
    section_match = re.search(
        r"Transaction summary(.*?)Charges and income summary",
        text,
        re.S
    )

    # ✅ Skip file if no trades
    if not section_match:
        return trades

    section = section_match.group(1)
    lines = [line.strip() for line in section.split("\n") if line.strip()]

    # ---- LOOP THROUGH LINES ----
    for i, line in enumerate(lines):

        if "Type:" in line:

            # ---- TYPE ----
            if "SAL" in line:
                trade_type = "Sell"
            elif "PUR" in line:
                trade_type = "Buy"
            else:
                trade_type = "Unknown"

            # ---- LOOK BACK ----
            if i < 2:
                continue

            data_line = lines[i-1]
            sec_line = lines[i-2]

            # ---- DATE ----
            date_match = re.search(r"(\d{2}[A-Z]{3}\d{4})", data_line)
            trade_date = format_date(date_match.group(1)) if date_match else None

            # ---- QUANTITY ----
            qty_match = re.search(
                r"\s(\d+\.?\d*)-?\s+USD|\s(\d+\.?\d*)-?\s+HKD",
                data_line
            )

            quantity = None
            if qty_match:
                quantity = qty_match.group(1) or qty_match.group(2)

            # ✅ APPEND (NOW INCLUDING ACCOUNT ID)
            trades.append({
                "Account Name": account_name,
                "Account ID": account_id,   # ✅ NEW COLUMN
                "Trade Date": trade_date,
                "Type": trade_type,
                "Security": sec_line,
                "Quantity": quantity
            })

    return trades


# ---------- MAIN LOOP ----------
all_trades = []

for filename in os.listdir(folder_path):

    if filename.lower().endswith(".pdf"):

        file_path = os.path.join(folder_path, filename)

        print(f"Processing: {filename}")

        file_trades = extract_trades(file_path)

        if file_trades:
            all_trades.extend(file_trades)
        else:
            print(f"No trades found in: {filename}")


# ---------- EXPORT ----------
df = pd.DataFrame(all_trades)

# ✅ Sort by date
df["Trade Date"] = pd.to_datetime(df["Trade Date"], format="%d-%m-%Y", errors="coerce")
df = df.sort_values(by="Trade Date")

# Convert back to string
df["Trade Date"] = df["Trade Date"].dt.strftime("%d-%m-%Y")

# ✅ Save
df.to_excel(output_file, index=False)

print("\n✅ DONE — All trades exported to:", output_file)
print(df)


# # Combine all records from different brokers

# In[38]:


import pandas as pd
import os
import re
from datetime import datetime

# ✅ Broker file paths
files = {
    "HSBC": r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\HSBC_Files\HSBC_trades.xlsx",
    "FUTU": r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Futu_Files\Futu_trades.xlsx",
    "IBKR": r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\IBKR_Files\IBKR_trades.xlsx"
}

# ✅ Output file
output_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Combined_Broker_Trades.xlsx"


# ✅ Function: detect Chinese characters
def contains_chinese(text):
    if isinstance(text, str):
        return bool(re.search(r'[\u4e00-\u9fff]', text))
    return False


all_dfs = []

# ---- LOOP EACH BROKER ----
for broker, file_path in files.items():

    print(f"Processing: {broker}")

    try:
        df = pd.read_excel(file_path, dtype={"Account ID": str})

        # ✅ FUTU handling
        if broker == "FUTU" and "Securities (Eng)" in df.columns:

            # ✅ Preserve original
            original_security = df["Security"]

            # ✅ Assign English to Security
            df["Security"] = df["Securities (Eng)"]

            # ✅ Only assign Chinese if it contains Chinese characters
            df["Security(cn)"] = original_security.apply(
                lambda x: x if contains_chinese(x) else None
            )

        else:
            # ✅ Other brokers
            df["Security(cn)"] = None

        # ✅ Add Broker
        df["Broker"] = broker

        # ✅ Keep only required columns
        df = df[
            ["Account Name", "Account ID","Trade Date", "Type", "Security", "Security(cn)", "Quantity", "Broker"]
        ]

        all_dfs.append(df)

    except Exception as e:
        print(f"Error reading {broker}: {e}")


# ---- COMBINE ----
combined_df = pd.concat(all_dfs, ignore_index=True)


# ✅ Standardize date
combined_df["Trade Date"] = pd.to_datetime(
    combined_df["Trade Date"], format="%d-%m-%Y", errors="coerce"
)

# ✅ Sort
combined_df = combined_df.sort_values(by=["Trade Date", "Account Name"])

# ✅ Convert back to string
combined_df["Trade Date"] = combined_df["Trade Date"].dt.strftime("%d-%m-%Y")


# ✅ Save
combined_df.to_excel(output_file, index=False)


print("\n✅ DONE — Combined file saved to:")
print(output_file)
print(combined_df)


# In[ ]:


# add staff name column
import pandas as pd

# ✅ Combined file (your output)
combined_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Combined_Broker_Trades.xlsx"

# ✅ Mapping file
mapping_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Staff Name and Account Name.xlsx"

# ✅ Output file
output_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Combined_Broker_Trades.xlsx"


# ---- LOAD FILES ----
df = pd.read_excel(combined_file)
mapping = pd.read_excel(mapping_file)


# ✅ Clean column names (very important)
mapping.columns = mapping.columns.str.strip()
df.columns = df.columns.str.strip()


# ---- CREATE MAPPING DICTIONARY ----
# Key: Statements name → Value: CIGP name
name_map = dict(zip(
    mapping["Staff Names as per Statements"],
    mapping["Staff Names as per CIGP"]
))


# ---- MAP STAFF NAME ----
df["Staff Name"] = df["Account Name"].map(name_map)


# ✅ If no match → keep original (optional but recommended)
df["Staff Name"] = df["Staff Name"].fillna(df["Account Name"])


# ---- REORDER COLUMNS (put Staff Name next to Account Name) ----
cols = df.columns.tolist()

# Remove Staff Name and reinsert after Account Name
cols.remove("Staff Name")
account_index = cols.index("Account Name")
cols.insert(account_index + 1, "Staff Name")

df = df[cols]


# ---- SAVE OUTPUT ----
df.to_excel(output_file, index=False)


print("\n✅ DONE — Staff Name added successfully")
print("Saved to:", output_file)
print(df)


# In[39]:


# add staff name column
import pandas as pd
from rapidfuzz import fuzz


# =========================
# ✅ Files
# =========================
combined_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Combined_Broker_Trades.xlsx"
mapping_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Staff Name and Account Name.xlsx"
output_file = combined_file


# =========================
# ✅ Load
# =========================
df = pd.read_excel(combined_file)
mapping = pd.read_excel(mapping_file)

mapping.columns = mapping.columns.str.strip()
df.columns = df.columns.str.strip()


# =========================
# ✅ Normalize function
# =========================
def normalize(text):
    return (
        str(text)
        .lower()
        .replace(" ", "")
        .replace(".", "")
        .replace(",", "")
    )


# =========================
# ✅ Build mapping list
# =========================
mapping_list = list(zip(
    mapping["Staff Names as per Statements"],
    mapping["Staff Names as per CIGP"]
))


# =========================
# ✅ Improved matching function
# =========================
def match_staff(account_name):

    account_norm = normalize(account_name)

    best_score = 0
    best_staff = None

    for stmt_name, cigp_name in mapping_list:

        stmt_norm = normalize(stmt_name)

        score = fuzz.ratio(account_norm, stmt_norm)

        if score > best_score:
            best_score = score
            best_staff = cigp_name

    # ✅ threshold to avoid wrong match
    if best_score >= 80:
        return best_staff
    else:
        return account_name   # fallback


# =========================
# ✅ Apply matching
# =========================
df["Staff Name"] = df["Account Name"].apply(match_staff)


# =========================
# ✅ Reorder columns
# =========================
cols = df.columns.tolist()

cols.remove("Staff Name")
account_index = cols.index("Account Name")
cols.insert(account_index + 1, "Staff Name")

df = df[cols]


# =========================
# ✅ Save
# =========================
df.to_excel(output_file, index=False)


print("\n✅ DONE — Robust staff matching applied")
print("Saved to:", output_file)


# In[3]:


## mtach ISIN
import requests

FIGI_API_KEY = ""  # optional


def get_isin_from_figi(security):

    url = "https://api.openfigi.com/v3/search"

    headers = {
        "Content-Type": "application/json"
    }

    if FIGI_API_KEY:
        headers["X-OPENFIGI-APIKEY"] = FIGI_API_KEY

    payload = [{
        "query": security
    }]

    try:
        response = requests.post(url, json=payload, headers=headers)
        data = response.json()

        if data and "data" in data[0]:
            return data[0]["data"][0].get("isin", None)

    except:
        pass

    return None


# In[10]:


import pandas as pd
import requests
import time


file_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Combined_Broker_Trades.xlsx"
df = pd.read_excel(file_path)


# =========================
# ✅ Extract ticker
# =========================
def extract_ticker(text):

    if pd.isna(text):
        return None

    for w in str(text).split():
        if w.isalpha() and w.isupper() and 2 <= len(w) <= 5:
            if w not in ["SHS", "PRE", "USD", "ORD"]:
                return w

    return None


# =========================
# ✅ OpenFIGI correct API
# =========================
def get_isin_from_figi(ticker):

    url = "https://api.openfigi.com/v3/mapping"

    headers = {
        "Content-Type": "application/json"
        # 👉 Add API Key if you have:
        # "X-OPENFIGI-APIKEY": "YOUR_KEY"
    }

    payload = [{
        "idType": "TICKER",
        "idValue": ticker
    }]

    try:
        response = requests.post(url, json=payload, headers=headers)

        print(f"\nTicker: {ticker}")
        print("Status:", response.status_code)

        if response.status_code != 200:
            print("❌ API error")
            return None

        data = response.json()
        print("Response:", data)

        if data and "data" in data[0]:
            isin = data[0]["data"][0].get("isin", None)
            print("✅ ISIN:", isin)
            return isin

    except Exception as e:
        print("❌ Exception:", e)

    return None

def get_isin_from_name(security):

    url = "https://api.openfigi.com/v3/search"   # ✅ correct for name search

    headers = {
        "Content-Type": "application/json"
        # add API key if you have
    }

    payload = [{
        "query": security
    }]

    try:
        r = requests.post(url, json=payload, headers=headers)

        print("\nSecurity:", security)
        print("Status:", r.status_code)

        if r.status_code != 200:
            return None

        data = r.json()
        print("Response:", data)

        # ✅ pick best match
        if isinstance(data, list) and "data" in data[0]:
            for item in data[0]["data"]:

                # ✅ prefer equity / ETF
                if item.get("marketSector") == "Equity":
                    return item.get("isin", None)

    except Exception as e:
        print("Error:", e)

    return None

# =========================
# ✅ MAIN LOOP
# =========================
isin_list = []

for _, row in df.iterrows():

    security = row["Security"]

    isin = get_isin_from_name(security)

    time.sleep(1)  # avoid API limit

    isin_list.append(isin)

df["ISIN"] = isin_list

# =========================
# ✅ SAVE
# =========================
output_path = file_path.replace(".xlsx", "_with_ISIN.xlsx")
df.to_excel(output_path, index=False)

print("\n✅ DONE — ISIN column generated")


# # Match with PAD protal records

# In[42]:


import pandas as pd
from rapidfuzz import fuzz
from pandas.tseries.offsets import BDay
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import yfinance as yf
import re


# =========================
# ✅ Load data
# =========================
trade_df = pd.read_excel(r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Combined_Broker_Trades.xlsx")
pad_df = pd.read_excel(r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Export from PAD Portal.xlsx")

trade_df["Trade Date"] = pd.to_datetime(trade_df["Trade Date"], dayfirst=True)
pad_df["Request Date"] = pd.to_datetime(pad_df["Request Date"], dayfirst=True)


# =========================
# ✅ Normalize + enhance
# =========================
def normalize(text):
    text = str(text).lower()
    remove_words = ["limited","ltd","inc","corp","holdings","group","class","co","shs"]
    for w in remove_words:
        text = text.replace(w,"")
    return " ".join(text.replace("-", " ").split())


def enhance_name(text):
    text = normalize(text)
    replacements = {
        "jingdong": "jd",
        "jian kang": "health",
        "qi che": "auto",
        "lanyueliang": "blue moon",
        "lan yue liang": "blue moon",
        "mei tuan": "meituan"
    }
    for k,v in replacements.items():
        text = text.replace(k,v)
    return text


# =========================
# ✅ ETF detection (IMPROVED)
# =========================
def is_etf(security):
    sec = str(security).lower()
    return (
        "etf" in sec or
        "index" in sec or
        "nasdaq" in sec or
        "track" in sec or
        "pre" in sec
    )


# ✅ Better ticker extraction
def extract_ticker(text):

    words = str(text).split()

    for w in words:
        clean = re.sub(r'[^A-Z0-9]', '', w)

        if (
            clean.isalnum() and
            clean.isupper() and
            2 <= len(clean) <= 6 and
            clean not in ["SHS", "USD", "HKD", "ORD"]
        ):
            return clean

    return None

# ✅ ETF AUM cache
ETF_CACHE = {}

def get_etf_aum(ticker):

    if ticker in ETF_CACHE:
        return ETF_CACHE[ticker]

    try:
        info = yf.Ticker(ticker).info
        aum = info.get("totalAssets", None)
        ETF_CACHE[ticker] = aum
        return aum
    except:
        ETF_CACHE[ticker] = None
        return None


# ✅ ETF status
def check_etf_status(security):

    sec = str(security)

    # ✅ Step 1: Known ETF keywords (STRICT)
    ETF_KEYWORDS = ["ETF"]

    if any(k in sec.upper() for k in ETF_KEYWORDS):
        keyword_flag = True
    else:
        keyword_flag = False

    # ✅ Step 2: Extract ticker safely
    ticker = extract_ticker(sec)

    if not ticker:
        return "NOT_ETF", None, None

    try:
        info = yf.Ticker(ticker).info
        quote_type = info.get("quoteType", None)

        # ✅ Step 3: Only trust Yahoo ETF label
        if quote_type != "ETF":
            return "NOT_ETF", ticker, None

        # ✅ Step 4: AUM check
        aum = info.get("totalAssets", None)

        if aum is None:
            return "ETF_UNCERTAIN", ticker, None

        if aum > 200_000_000:
            return "ETF_EXEMPT", ticker, aum
        else:
            return "ETF_CONTROL", ticker, aum

    except:
        # ✅ ONLY fallback if strong keyword signal exists
        if keyword_flag:
            return "ETF_UNCERTAIN", ticker, None
        else:
            return "NOT_ETF", ticker, None


# =========================
# ✅ Late trade rule
# =========================
def is_late_trade(req_date, trade_date):
    return trade_date > req_date + BDay(3)


# =========================
# ✅ Security selection
# =========================
def get_trade_security(row):
    if "Securities (Eng)" in trade_df.columns:
        return row["Securities (Eng)"]
    return row["Security"]


# =========================
# ✅ MAIN PROCESS
# =========================
results = []

for _, trade in trade_df.iterrows():

    staff = trade["Staff Name"]
    trade_date = trade["Trade Date"]
    trade_sec = get_trade_security(trade)
    trade_type = trade["Type"]
    trade_qty = trade["Quantity"]

    issues = []

    # ✅ ETF check
    etf_status, ticker, aum = check_etf_status(trade_sec)

    if etf_status == "ETF_EXEMPT":
        results.append({
            **trade,
            "Matched Security": "ETF Exempt",
            "Match Score": None,
            "Request Date": None,
            "Approved Qty": None,
            "ETF Status": etf_status,
            "AUM": aum,
            "Ticker": ticker,
            "Unmatched Approval Security": "",
            "Unmatched Approval Security Request Qty": "",
            "Issue": "OK (ETF >200M Exempt)"
        })
        continue

    if etf_status == "ETF_UNCERTAIN":
        issues.append("ETF AUM Uncertain")

    # ✅ PAD matching
    pad_filtered = pad_df[pad_df["Requestor"] == staff].copy()
    pad_filtered = pad_filtered[pad_filtered["Request Date"] <= trade_date]

    if pad_filtered.empty:
        results.append({**trade, "Issue": "No prior PAD approval"})
        continue

    pad_filtered["Date Diff"] = (trade_date - pad_filtered["Request Date"]).dt.days
    pad_filtered = pad_filtered.sort_values("Date Diff").head(5)

    best_score = 0
    best_row = None

    for _, p in pad_filtered.iterrows():
        score = fuzz.token_sort_ratio(
            enhance_name(trade_sec),
            enhance_name(p["Security"])
        )

        if score > best_score:
            best_score = score
            best_row = p

    if best_score < 70:
        issues.append("Low Match Score")

    if best_score < 50 or best_row is None:
        issues.append("Unapproved Security")

        results.append({
            **trade,
            "Matched Security": None,
            "Match Score": best_score,
            "ETF Status": etf_status,
            "AUM": aum,
            "Ticker": ticker,
            "Unmatched Approval Security": "",
            "Unmatched Approval Security Request Qty": "",
            "Issue": "; ".join(issues)
        })
        continue

    req_date = best_row["Request Date"]
    req_type = best_row["Type"]
    req_qty = best_row["Quantity"]

    if is_late_trade(req_date, trade_date):
        issues.append("Late Trade")

    if str(trade_type).lower() != str(req_type).lower():
        issues.append("Type Mismatch")

    if trade_qty > req_qty:
        issues.append("Quantity Exceeded")

    results.append({
        **trade,
        "Matched Security": best_row["Security"],
        "Match Score": best_score,
        "Request Date": req_date,
        "Approved Qty": req_qty,
        "ETF Status": etf_status,
        "AUM": aum,
        "Ticker": ticker,
        "Unmatched Approval Security": "",
        "Unmatched Approval Security Request Qty": "",
        "Issue": "OK" if not issues else "; ".join(issues)
    })


df = pd.DataFrame(results)


# =========================
# ✅ Unmatched approvals
# =========================
for key, group in df.groupby(["Staff Name", "Trade Date"]):

    staff, date = key

    pad_group = pad_df[
        (pad_df["Requestor"] == staff) &
        (pad_df["Request Date"] == date)
    ]

    pad_dict = dict(zip(pad_group["Security"], pad_group["Quantity"]))
    matched = set(group["Matched Security"].dropna())

    unmatched = set(pad_dict.keys()) - matched

    df.loc[group.index, "Unmatched Approval Security"] = ", ".join(unmatched)
    df.loc[group.index, "Unmatched Approval Security Request Qty"] = ", ".join(
        f"{k}: {pad_dict[k]}" for k in unmatched
    )

# =========================
# ✅ Format dates (REMOVE TIME)
# =========================
if "Trade Date" in df.columns:
    df["Trade Date"] = df["Trade Date"].dt.strftime("%d-%m-%Y")

if "Request Date" in df.columns:
    df["Request Date"] = pd.to_datetime(df["Request Date"], errors="coerce").dt.strftime("%d-%m-%Y")
# =========================
# ✅ SAVE
# =========================
df = df.drop(columns=["AUM", "Ticker", "ETF Status"], errors="ignore")
output_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\PAD_Control_Result.xlsx"
df.to_excel(output_path, index=False)


# =========================
# ✅ Highlight
# =========================
wb = load_workbook(output_path)
ws = wb.active

red = PatternFill(start_color="FF9999", fill_type="solid")
yellow = PatternFill(start_color="FFFF99", fill_type="solid")

issue_col = [c.value for c in ws[1]].index("Issue") + 1

for r in range(2, ws.max_row+1):
    val = ws.cell(r, issue_col).value

    if val:
        if "Unapproved Security" in val:
            for c in range(1, ws.max_column+1):
                ws.cell(r, c).fill = red

        elif "ETF AUM Uncertain" in val:
            for c in range(1, ws.max_column+1):
                ws.cell(r, c).fill = yellow


wb.save(output_path)

print("✅ FINAL VERSION COMPLETE")


# # For submission records check

# In[45]:


import pandas as pd

# ✅ Input file
file_path = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\PAD_Control_Result.xlsx"

# ✅ Output file
output_file = r"C:\Users\KateZhang\OneDrive - CIGP SA\Desktop\PAD Automation\Statement_Control_Records.xlsx"


# ====================
# ✅ LOAD DATA
# ====================
df = pd.read_excel(file_path)

# ✅ Convert date
df["Trade Date"] = pd.to_datetime(
    df["Trade Date"],
    format="%d-%m-%Y",
    errors="coerce"
)


# ✅ ✅ FIX 1: REMOVE INVALID DATES (avoid NaT column)
df = df.dropna(subset=["Trade Date"])

# ✅ Extract Year-Month
df["Year-Month"] = df["Trade Date"].dt.to_period("M").astype(str)


# ====================
# ✅ CREATE BASE TABLE (WITH ACCOUNT ID ✅)
# ====================
base = df[["Staff Name", "Account Name", "Account ID"]].drop_duplicates()


# ====================
# ✅ CREATE MONTH LIST (NO NaT ✅)
# ====================
all_months = sorted(df["Year-Month"].dropna().unique())


# ====================
# ✅ BUILD CONTROL MATRIX
# ====================
result = base.copy()

for month in all_months:

    temp = df[df["Year-Month"] == month][["Account ID"]].drop_duplicates()

    result[month] = result["Account ID"].isin(temp["Account ID"]).map({
        True: "✅",
        False: "❌"
    })


# ====================
# ✅ SORT
# ====================
result = result.sort_values(by=["Staff Name", "Account Name"])


# ====================
# ✅ SAVE
# ====================
result.to_excel(output_file, index=False)

print("✅ DONE — Clean control file created")
print("Saved to:", output_file)
print(result)


# In[ ]:


print("Version 2 - test")

