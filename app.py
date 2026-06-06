# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import re
from datetime import datetime, time
from io import BytesIO

GRACE = 5
WEEKEND_DAYS = ["Friday", "Saturday"]
WEEKEND_START = time(8, 0)

NAME_MAP = {
    "Jomana": "jumana", "Jomanaa": "jumana", "Huda": "Hoda", "Roaa": "Roua",
    "Njoud": "Nujood", "yoni": "Yuni", "Amir": "Ameer", "yazeed": "Yazeed",
    "noor": "Noor", "sraa": "Sraa", "Riri": "RiRi", "Jowana": "Jowana",
    "Kareem": "Kareem", "Lujain": "Lujain", "Rajaa": "Rajaa",
    "Ashwaq": "Ashwaq", "Ehda": "Ehdaa", "Hasmena": "Hasmena",
    "Raad": "Raad", "Saeed": "Saeed", "Malek": "Malek",
    "Abdulmalik": "Abdulmalik", "Shata": "Shata", "Noor": "Noor",
    "Sraa": "Sraa", "Yoni": "Yuni", "Yazeed": "Yazeed", "Razan": "Razan",
    "Jumana": "jumana", "jasmin": "jasmin", "Nagat": "Nagat",
}

WEEKDAY_AR = {"Monday": "الاثنين", "Tuesday": "الثلاثاء", "Wednesday": "الأربعاء",
              "Thursday": "الخميس", "Friday": "الجمعة", "Saturday": "السبت", "Sunday": "الأحد"}

MONTH_NAMES_EN = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}

st.set_page_config(page_title="Meraki - حساب التأخير", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    * { font-family: 'Cairo', sans-serif !important; }
</style>
""", unsafe_allow_html=True)

def parse_time_start(time_str, col_index=None):
    if pd.isna(time_str): return None
    time_str = str(time_str).replace("\n", " ").strip()
    for pattern in [r'(\d{1,2}):(\d{2})\s*[-~]\s*\d{1,2}:\d{2}', r'(\d{1,2}):(\d{2})\s*-*\s*']:
        m = re.match(pattern, time_str)
        if m:
            hour, minute = int(m.group(1)), int(m.group(2))
            if col_index is not None and col_index >= 8:
                if 1 <= hour <= 11: hour += 12
            else:
                if 1 <= hour <= 5: hour += 12
            return time(hour, minute)
    return None

def clean_name(name):
    name = str(name).strip()
    name = re.sub(r'\s*\d{1,2}:\d{2}\s*$', '', name)
    name = re.sub(r'\s+\d+$', '', name)
    return name.strip()

def is_valid_name(name):
    name = str(name).strip()
    if len(name) <= 1: return False
    skip = ["artists", "hours", "comments", "days", "adjust days", "nan", "off",
            "part time", "training", "absent", "sickleave", "team meeting",
            "بلديه", "بادية", "بلدية", "تبديل", "استئذان", "eid", "vicaiton",
            "vication", "vacation", "weekend", "holiday"]
    if any(k in name.lower() for k in skip): return False
    if name.replace(".", "").replace("-", "").replace(":", "").replace(" ", "").isdigit(): return False
    if re.match(r'^\d{1,2}:\d{2}$', name): return False
    return True

def fmt12(t):
    if t is None: return "---"
    h, m = t.hour, t.minute
    if h == 0: return f"12:{m:02d} AM"
    elif h == 12: return f"12:{m:02d} PM"
    elif h > 12: return f"{h-12}:{m:02d} PM"
    else: return f"{h}:{m:02d} AM"

def extract_schedule(df_sheet):
    headers = df_sheet.iloc[1].tolist()
    col_to_time = {}
    for i, h in enumerate(headers):
        t = parse_time_start(h, col_index=i)
        if t: col_to_time[i] = t
    schedule = {}
    cur_date = None
    for idx in range(2, len(df_sheet)):
        row = df_sheet.iloc[idx]
        if pd.notna(row[0]):
            try:
                d = row[0]
                cur_date = d.date() if isinstance(d, (datetime, pd.Timestamp)) else pd.to_datetime(d).date()
                schedule[cur_date] = {}
            except:
                cur_date = None
                continue
        if cur_date is None: continue
        if len(row) > 36 and pd.notna(row[36]):
            if str(row[36]).strip().lower() in ["artists", "hours", "comments", "days", "adjust days"]:
                continue
        is_weekend = cur_date.strftime("%A") in WEEKEND_DAYS
        for col, start_time in col_to_time.items():
            if col < len(row) and pd.notna(row[col]):
                emp = clean_name(str(row[col]))
                if not is_valid_name(emp): continue
                actual_time = WEEKEND_START if (is_weekend and col < 8) else start_time
                if emp not in schedule[cur_date] or col < schedule[cur_date][emp][0]:
                    schedule[cur_date][emp] = (col, actual_time)
    for d in schedule:
        schedule[d] = {emp: t for emp, (col, t) in schedule[d].items()}
    return schedule

def read_schedule(excel_file):
    xls = pd.ExcelFile(excel_file)
    available = xls.sheet_names
    sheet_names = [s for s in available if any(m in s.lower() for m in ['may','jun','jul','aug','sep','oct','nov','dec','jan','feb','mar','apr'])] or available
    all_sched = {}
    for sheet in sheet_names:
        if sheet not in available: continue
        df = pd.read_excel(excel_file, sheet_name=sheet, header=None)
        sched = extract_schedule(df)
        for d, emps in sched.items():
            all_sched.setdefault(d, {}).update(emps)
    return {d: e for d, e in all_sched.items() if e}

def parse_attendance_excel(att_file):
    df = pd.read_excel(att_file)
    df.columns = [str(c).strip() for c in df.columns]
    name_col = next((c for c in df.columns if c.lower() in ['name', 'employee']), df.columns[2] if len(df.columns) > 2 else df.columns[1])
    time_col = next((c for c in df.columns if c.lower() in ['time', 'datetime']), df.columns[3] if len(df.columns) > 3 else df.columns[2])
    state_col = next((c for c in df.columns if c.lower() in ['state', 'status']), df.columns[4] if len(df.columns) > 4 else df.columns[3])
    records = []
    for _, row in df.iterrows():
        try:
            name, time_val, state = str(row[name_col]).strip(), row[time_col], str(row[state_col]).strip()
            if state not in ['C/In', 'C/Out', 'In', 'Out'] or pd.isna(time_val): continue
            dt = pd.to_datetime(time_val)
            records.append({'date': dt.date(), 'pdf_name': name, 'cin_time': dt.time(),
                           'type': 'C/In' if state in ['C/In', 'In'] else 'C/Out'})
        except: continue
    cin_only = [r for r in records if r['type'] == 'C/In']
    first_cin = {}
    for r in cin_only:
        key = (r['date'], r['pdf_name'])
        if key not in first_cin or r['cin_time'] < first_cin[key]['cin_time']:
            first_cin[key] = r
    df_out = pd.DataFrame([{"date": k[0], "pdf_name": k[1], "cin_time": v["cin_time"]}
                           for k, v in first_cin.items()])
    reverse_map = {v: k for k, v in NAME_MAP.items()}
    df_out["employee"] = df_out["pdf_name"].map(reverse_map).fillna(df_out["pdf_name"])
    return df_out

def calc(schedule, att_df):
    results = []
    for date, employees in sorted(schedule.items()):
        for emp, sched_time in sorted(employees.items()):
            is_evening = (sched_time.hour >= 12)
            day = att_df[(att_df["date"] == date) & (att_df["employee"] == emp)]
            rec = day[day["cin_time"] >= time(12, 0)] if is_evening else day[day["cin_time"] < time(12, 0)]
            if len(rec) == 0 and len(day) > 0: rec = day
            month_name = f"{MONTH_NAMES_EN[date.month]} {date.year}"
            if len(rec) == 0:
                results.append({"التاريخ": date.strftime("%d/%m/%Y"), "اليوم": WEEKDAY_AR[date.strftime("%A")],
                    "الشهر": month_name, "الموظف": emp, "وقت_الدوام": fmt12(sched_time),
                    "وقت_البصمة": "---", "الفرق_دقيقة": None, "التأخير": None, "الحالة": "لا توجد بصمة"})
                continue
            cin = rec.iloc[0]["cin_time"]
            sm = sched_time.hour * 60 + sched_time.minute
            cm = cin.hour * 60 + cin.minute
            diff = cm - sm
            if diff > 240: status, delay = "وردية مسائية", 0
            elif diff > GRACE: status, delay = "متأخر", diff - GRACE
            else: status, delay = "في الوقت", 0
            results.append({"التاريخ": date.strftime("%d/%m/%Y"), "اليوم": WEEKDAY_AR[date.strftime("%A")],
                "الشهر": month_name, "الموظف": emp, "وقت_الدوام": fmt12(sched_time),
                "وقت_البصمة": fmt12(cin), "الفرق_دقيقة": diff, "التأخير": delay, "الحالة": status})
    return pd.DataFrame(results)

def save_report(df):
    df = df.sort_values(["الشهر", "التاريخ", "الموظف"])
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="الكل", index=False)
        for month in sorted(df["الشهر"].unique()):
            mdf = df[df["الشهر"] == month]
            sn = month.replace(" ", "_")[:31]
            mdf.to_excel(writer, sheet_name=sn, index=False)
            summary = mdf.groupby("الموظف").agg(
                ايام_العمل=("التاريخ", "count"),
                ايام_متأخر=("الحالة", lambda x: (x == "متأخر").sum()),
                ايام_في_الوقت=("الحالة", lambda x: (x == "في الوقت").sum()),
                ايام_بدون_بصمة=("الحالة", lambda x: (x == "لا توجد بصمة").sum()),
                مجموع_التأخير=("التأخير", lambda x: x.fillna(0).sum()),
            )
            summary["متوسط_التأخير"] = (summary["مجموع_التأخير"] / summary["ايام_العمل"]).round(2)
            summary.sort_values("مجموع_التأخير", ascending=False).to_excel(writer, sheet_name=f"ملخص_{sn[:20]}")
            late = mdf[mdf["الحالة"] == "متأخر"]
            if len(late) > 0: late.to_excel(writer, sheet_name=f"تأخير_{sn[:20]}", index=False)
    output.seek(0)
    return output, df

st.markdown("""
<div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    padding: 30px; border-radius: 20px; text-align: center; margin-bottom: 30px;
    box-shadow: 0 10px 40px rgba(0,0,0,0.3);">
    <h1 style="color: #e94560; font-size: 2.5rem; font-weight: 700; margin: 0;">🍽️ MERAKI</h1>
    <p style="color: #eee; font-size: 1.1rem; margin: 10px 0 0 0;">نظام حساب التأخير — بسيط وسريع</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("### 📅 جدول الدوامات")
    sched_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"], key="sched")
with col2:
    st.markdown("### 📝 ملف البصمات")
    att_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"], key="att")

if sched_file and att_file:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 ابدأ الحساب", use_container_width=True):
        with st.spinner("⏳ جاري المعالجة..."):
            try:
                schedule = read_schedule(sched_file)
                att_df = parse_attendance_excel(att_file)
                att_df = att_df.drop_duplicates(subset=["date", "employee"], keep="first")
                results = calc(schedule, att_df)
                excel_file, df_display = save_report(results)
                st.success("✅ تم الحساب بنجاح!")
                total = len(results)
                late = len(results[results["الحالة"] == "متأخر"])
                ontime = len(results[results["الحالة"] == "في الوقت"])
                missing = len(results[results["الحالة"] == "لا توجد بصمة"])
                c1, c2, c3, c4 = st.columns(4)
                c1.markdown(f'<div style="background: #16213e; border-radius: 15px; padding: 20px; text-align: center;"><h3 style="color: #e94560; font-size: 2rem; margin: 0;">{total}</h3><p style="color: #aaa;">إجمالي الأيام</p></div>', unsafe_allow_html=True)
                c2.markdown(f'<div style="background: #16213e; border-radius: 15px; padding: 20px; text-align: center;"><h3 style="color: #00d9ff; font-size: 2rem; margin: 0;">{ontime}</h3><p style="color: #aaa;">✅ في الوقت</p></div>', unsafe_allow_html=True)
                c3.markdown(f'<div style="background: #16213e; border-radius: 15px; padding: 20px; text-align: center;"><h3 style="color: #e94560; font-size: 2rem; margin: 0;">{late}</h3><p style="color: #aaa;">⚠️ متأخر</p></div>', unsafe_allow_html=True)
                c4.markdown(f'<div style="background: #16213e; border-radius: 15px; padding: 20px; text-align: center;"><h3 style="color: #ffd700; font-size: 2rem; margin: 0;">{missing}</h3><p style="color: #aaa;">❌ بدون بصمة</p></div>', unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("### 📊 التفاصيل")
                st.dataframe(df_display, use_container_width=True, height=500)
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 تحميل التقرير (Excel)",
                    data=excel_file,
                    file_name=f"تقرير_التأخير_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ حدث خطأ: {str(e)}")
                st.info("💡 تأكد من صحة الملفات")
else:
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("""
    ### 📌 كيف الاستخدام:
    1. **ارفع ملف الجدول** (Excel) — ملف الدوامات الشهرية
    2. **ارفع ملف البصمات** (Excel) — من جهاز البصمة
    3. **اضغط ابدأ الحساب** — والتقرير يجهز تلقائياً!
    """)
