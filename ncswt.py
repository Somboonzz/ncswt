import streamlit as st
import pandas as pd
import altair as alt
import datetime
import os
import pytz 

# ----------------------------------------------------------------------------------
# ตั้งค่าหน้า และ CSS
# ----------------------------------------------------------------------------------
st.set_page_config(page_title="HR Dashboard", layout="wide")

st.markdown("""
    <style>
        /* CSS สำหรับตารางและอื่น ๆ (ตามโค้ดเดิม) */
        div[data-testid="stDataframeHeader"] div {
            text-align: center !important;
            vertical-align: middle !important;
            justify-content: center !important;
        }
        
        div[data-testid="stDataframeCell"] {
            text-align: center !important;
            justify-content: center !important;
        }

        .stDataFrame {
            margin-left: 1rem;
            margin-right: 1rem;
        }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------------------------------------
# โซนเวลาไทย และฟังก์ชันช่วย
# ----------------------------------------------------------------------------------
bangkok_tz = pytz.timezone("Asia/Bangkok")

def thai_date(dt):
    return dt.strftime(f"%d/%m/{dt.year + 543}")

thai_months = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def format_thai_month(period):
    year = period.year + 543
    month = thai_months[period.month - 1]
    return f"{month} {year}"

# ----------------------------------------------------------------------------------
# โหลดไฟล์ Excel
# ----------------------------------------------------------------------------------
@st.cache_data(ttl=300)
def load_data(file_path="attendances.xlsx"):
    try:
        if file_path and os.path.exists(file_path):
            df = pd.read_excel(file_path, engine='openpyxl', dtype={'เข้างาน': str, 'ออกงาน': str})
            return df
        else:
            st.warning("❌ ไม่พบไฟล์ Excel: attendances.xlsx")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ อ่านไฟล์ Excel ไม่ได้: {e}")
        return pd.DataFrame()

df = load_data()

# ----------------------------------------------------------------------------------
# ปุ่ม Refresh และแสดงเวลา
# ----------------------------------------------------------------------------------
if st.button("🔄 Refresh ข้อมูล (Manual)"):
    st.cache_data.clear()
    st.rerun()

bangkok_now = datetime.datetime.now(pytz.utc).astimezone(bangkok_tz)
st.markdown(
    f"<div style='text-align:right; font-size:50px; color:#FF5733; font-weight:bold;'>"
    f"🗓 {thai_date(bangkok_now)} | ⏰ {bangkok_now.strftime('%H:%M:%S')}</div>",
    unsafe_allow_html=True
)

# ----------------------------------------------------------------------------------
# Dashboard หลัก (การเตรียมข้อมูล)
# ----------------------------------------------------------------------------------
if not df.empty:
    for col in ["ชื่อ-สกุล", "แผนก", "ข้อยกเว้น"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    if "แผนก" in df.columns:
        df["แผนก"] = df["แผนก"].replace({"nan": "ไม่ระบุ", "": "ไม่ระบุ"})

    if "วันที่" in df.columns:
        df["วันที่"] = pd.to_datetime(df["วันที่"], errors='coerce')
        df["ปี"] = df["วันที่"].dt.year + 543
        df["เดือน"] = df["วันที่"].dt.to_period("M")
    
    if 'เข้างาน' in df.columns:
        df['เวลาเข้า'] = pd.to_datetime(df['เข้างาน'], format='%H:%M:%S', errors='coerce').dt.time
        df['เวลาเข้า'] = df['เวลาเข้า'].apply(lambda t: datetime.time(0, 0) if pd.isna(t) else t)
    if 'ออกงาน' in df.columns:
        df['เวลาออก'] = pd.to_datetime(df['ออกงาน'], format='%H:%M:%S', errors='coerce').dt.time
        df['เวลาออก'] = df['เวลาออก'].apply(lambda t: datetime.time(0, 0) if pd.isna(t) else t)

    df_filtered = df.copy()

    # --- Filter (ตามโค้ดเดิมของคุณ) ---
    col1, col2, col3 = st.columns(3)
    
    with col1:
        years = ["-- แสดงทั้งหมด --"] + sorted(df["ปี"].dropna().unique(), reverse=True)
        selected_year = st.selectbox("📆 เลือกปี", years)
        if selected_year != "-- แสดงทั้งหมด --":
            df_filtered = df_filtered[df_filtered["ปี"] == int(selected_year)]

    with col2:
        if "เดือน" in df_filtered.columns and not df_filtered.empty:
            available_months = sorted(df_filtered["เดือน"].dropna().unique())
            month_options = ["-- แสดงทั้งหมด --"] + [format_thai_month(m) for m in available_months]
            selected_month = st.selectbox("📅 เลือกเดือน", month_options)
            if selected_month != "-- แสดงทั้งหมด --":
                mapping = {format_thai_month(m): str(m) for m in available_months}
                selected_period = mapping[selected_month]
                df_filtered = df_filtered[df_filtered["เดือน"].astype(str) == selected_period]

    with col3:
        departments = ["-- แสดงทั้งหมด --"] + sorted(df_filtered["แผนก"].dropna().unique())
        selected_dept = st.selectbox("🏢 เลือกแผนก", departments)
        if selected_dept != "-- แสดงทั้งหมด --":
            df_filtered = df_filtered[df_filtered["แผนก"] == selected_dept]
    # -----------------------------------

    # --- คำนวณประเภทการลา ---
    def leave_days(row):
        if "ครึ่งวัน" in str(row):
            return 0.5
        return 1

    df_filtered["ลาป่วย/ลากิจ"] = df_filtered["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ลาป่วย", "ลากิจ", "ลาป่วยครึ่งวัน", "ลากิจครึ่งวัน"] else 0
    )
    df_filtered["ขาด"] = df_filtered["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ขาด", "ขาดครึ่งวัน"] else 0
    )
    df_filtered["สาย"] = df_filtered["ข้อยกเว้น"].apply(lambda x: 1 if str(x) == "สาย" else 0)

    leave_types = ["ลาป่วย/ลากิจ", "ขาด", "สาย"]
    summary = df_filtered.groupby(["ชื่อ-สกุล", "แผนก"])[leave_types].sum().reset_index()

    st.title("📊 แดชบอร์ดการลา / ขาด / สาย")

    # --- ตัวกรองพนักงาน ---
    if 'selected_employee' not in st.session_state:
        st.session_state.selected_employee = '-- แสดงทั้งหมด --'
    all_names = ["-- แสดงทั้งหมด --"] + sorted(summary["ชื่อ-สกุล"].unique())
    selected_employee = st.selectbox("🔍 ค้นหาชื่อพนักงาน", all_names, key='selected_employee')

    # --- สีกราฟ: ยืนยันตามภาพที่ต้องการ (ลาป่วย/ลากิจ: แดงเข้ม, ขาด: ส้มแดง, สาย: เหลือง) ---
    colors = {
        "ลาป่วย/ลากิจ": "#C70039", 
        "ขาด": "#FF5733", 
        "สาย": "#FFC300", 
    }

# ----------------------------------------------------------------------------------
# ส่วนที่ 2: Tabs จัดอันดับ (ตามโค้ดเดิมของคุณ)
# ----------------------------------------------------------------------------------
tabs = st.tabs(leave_types)
for t, leave in zip(tabs, leave_types):
    with t:
        st.subheader(f"🏆 จัดอันดับ {leave} (แยกรายบุคคล)")
        
        if selected_employee != "-- แสดงทั้งหมด --":
            summary_filtered = summary[summary["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
            person_data_full = df_filtered[df_filtered["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
        else:
            summary_filtered = summary.reset_index(drop=True)
            person_data_full = df_filtered.reset_index(drop=True)

        st.markdown("### 📌 สรุปข้อมูลรายบุคคล")
        summary_filtered_display = summary_filtered[summary_filtered[leave_types].sum(axis=1) > 0]
        st.dataframe(summary_filtered_display, use_container_width=True, hide_index=True)

        # ... (ส่วนแสดงรายละเอียดวันลา) ...

        ranking = summary[["ชื่อ-สกุล", "แผนก", leave]].sort_values(by=leave, ascending=False).reset_index(drop=True)
        ranking.insert(0, "อันดับ", range(1, len(ranking) + 1))
        ranking_display = ranking
        if selected_employee != "-- แสดงทั้งหมด --":
            ranking_display = ranking_display[ranking_display["ชื่อ-สกุล"] == selected_employee]
        
        ranking_display = ranking_display[ranking_display[leave] > 0] 
        
        st.markdown("### 🏅 ตารางอันดับ")
        st.dataframe(ranking_display.reset_index(drop=True), use_container_width=True, hide_index=True)


# ----------------------------------------------------------------------------------
# Pie Chart (แก้ไข radius ให้ข้อความเข้าใกล้วงกลมมากขึ้น)
# ----------------------------------------------------------------------------------
st.markdown("---")
st.subheader("🥧 สัดส่วนรวมการลา/ขาด/สาย (พร้อมชื่อ + เปอร์เซ็นต์)")

total_summary = summary[leave_types].sum().reset_index()
total_summary.columns = ['ประเภท', 'ยอดรวม']
total_summary = total_summary[total_summary['ยอดรวม'] > 0].reset_index(drop=True)

if total_summary['ยอดรวม'].sum() > 0:
    total = total_summary['ยอดรวม'].sum()
    total_summary['Percentage'] = (total_summary['ยอดรวม'] / total * 100).round(1)
    total_summary['label'] = total_summary.apply(lambda x: f"{x['ประเภท']} {x['Percentage']}%", axis=1)

    # 1. สร้าง base chart
    base = alt.Chart(total_summary).encode(
        theta=alt.Theta("ยอดรวม", stack=True),
        color=alt.Color(
            "ประเภท",
            scale=alt.Scale(domain=list(colors.keys()), range=list(colors.values()))
        ),
        order=alt.Order("ยอดรวม", sort="descending"), 
        tooltip=[
            "ประเภท",
            alt.Tooltip("ยอดรวม", format=".1f", title="จำนวน (วัน/ครั้ง)"),
            alt.Tooltip("Percentage", format=".1f", title="เปอร์เซ็นต์ (%)")
        ]
    )

    # 2. วงกลมหลัก
    pie = base.mark_arc(outerRadius=160, innerRadius=60)

    # 3. Text Label (ข้อความด้านนอก)
    # *** ปรับค่า radius จาก 250 เป็น 190 (หรือ 180) ***
    text_labels = base.mark_text(
        radius= 190,  # <<< ลดค่านี้เพื่อดึงข้อความให้ใกล้ขึ้น >>>
        size=20,
        fontWeight="bold",
    ).encode(
        text=alt.Text('label:N'),
        color=alt.value('black') 
    )

    # 4. ข้อความตรงกลางวงกลม (รวม 100%)
    center_text = alt.Chart(pd.DataFrame({'text': [f"รวม 100%"]})).mark_text(
        size=20, color='black', fontWeight='bold'
    ).encode(text='text:N')

    # รวมทุกส่วน: วงกลม + ข้อความภายนอก + ข้อความกลาง
    chart = pie + text_labels + center_text
    
    chart = chart.properties(
        width=400,
        height=400,
        title="สัดส่วนรวมการลา/ขาด/สาย"
    )

    st.altair_chart(chart, use_container_width=True)

    # ตารางสรุป Pie Chart
    st.dataframe(total_summary, use_container_width=True, hide_index=True)

else:
    st.info("ไม่พบข้อมูลการลา/ขาด/สายในช่วงเวลาที่เลือก")