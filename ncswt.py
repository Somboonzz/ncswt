import streamlit as st
import pandas as pd
import altair as alt
import datetime
import os
import pytz

st.set_page_config(page_title="HR Dashboard", layout="wide")

# -----------------------------
# โซนเวลาไทย
# -----------------------------
bangkok_tz = pytz.timezone("Asia/Bangkok")

def thai_date(dt):
    # เปลี่ยนปี พ.ศ. ในฟังก์ชัน thai_date ให้ถูกต้อง (ปี + 543)
    return dt.strftime(f"%d/%m/{dt.year + 543}")

thai_months = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def format_thai_month(period):
    year = period.year + 543
    month = thai_months[period.month - 1]
    return f"{month} {year}"

# ฟังก์ชันช่วยจัดรูปแบบตัวเลข (ใช้ซ้ำ)
def format_value(val, is_time=False):
    """จัดรูปแบบตัวเลข: ทศนิยม 1 ตำแหน่งสำหรับค่าที่ไม่ใช่ 0/จำนวนเต็ม, จำนวนเต็มสำหรับค่าที่เป็น 0/จำนวนเต็ม"""
    if is_time:
         return f"{val}" # สำหรับ 'สาย'
    if val == 0:
         return "0"
    # ใช้ .1f สำหรับทศนิยม 1 ตำแหน่ง ถ้าไม่ใช่จำนวนเต็ม
    return f"{val:.1f}" if val != int(val) else f"{int(val)}"
    
# -----------------------------
# โหลดไฟล์ Excel
# -----------------------------
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

# -----------------------------
# ปุ่ม Refresh
# -----------------------------
if st.button("🔄 Refresh ข้อมูล (Manual)"):
    st.cache_data.clear()
    st.rerun()

# -----------------------------
# นาฬิกา
# -----------------------------
bangkok_now = datetime.datetime.now(pytz.utc).astimezone(bangkok_tz)
st.markdown(
    f"<div style='text-align:right; font-size:50px; color:#FF5733; font-weight:bold;'>"
    f"🗓 {thai_date(bangkok_now)}  |  ⏰ {bangkok_now.strftime('%H:%M:%S')}</div>",
    unsafe_allow_html=True
)

# -----------------------------
# Dashboard
# -----------------------------
if not df.empty:
    # --- ทำความสะอาดข้อมูล
    for col in ["ชื่อ-สกุล", "แผนก", "ข้อยกเว้น"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    if "แผนก" in df.columns:
        df["แผนก"] = df["แผนก"].replace({"nan": "ไม่ระบุ", "": "ไม่ระบุ"})

    if "วันที่" in df.columns:
        df["วันที่"] = pd.to_datetime(df["วันที่"], errors='coerce')
        # ตรวจสอบการแปลงปี พ.ศ.
        df["ปี"] = df["วันที่"].dt.year + 543
        df["เดือน"] = df["วันที่"].dt.to_period("M")

    # --- เวลาเข้า/ออก
    for col, new_col in [("เข้างาน", "เวลาเข้า"), ("ออกงาน", "เวลาออก")]:
        if col in df.columns:
            df[new_col] = pd.to_datetime(df[col], format='%H:%M:%S', errors='coerce').dt.time
            df[new_col] = df[new_col].apply(lambda t: datetime.time(0, 0) if pd.isna(t) else t)

    df_filtered = df.copy()

    # --- Filter ปี
    col1, col2, col3 = st.columns(3)
    with col1:
        years = ["-- แสดงทั้งหมด --"] + sorted(df["ปี"].dropna().unique(), reverse=True)
        selected_year = st.selectbox("📆 เลือกปี", years)
        if selected_year != "-- แสดงทั้งหมด --":
            df_filtered = df_filtered[df_filtered["ปี"] == int(selected_year)]

    # --- Filter เดือน
    with col2:
        if "เดือน" in df_filtered.columns and not df_filtered.empty:
            available_months = sorted(df_filtered["เดือน"].dropna().unique())
            month_options = ["-- แสดงทั้งหมด --"] + [format_thai_month(m) for m in available_months]
            selected_month = st.selectbox("📅 เลือกเดือน", month_options)
            if selected_month != "-- แสดงทั้งหมด --":
                # แปลงกลับจากชื่อเดือนไทยเป็น Period
                mapping = {format_thai_month(m): m for m in available_months}
                selected_period = mapping[selected_month]
                df_filtered = df_filtered[df_filtered["เดือน"] == selected_period]
        else:
            st.selectbox("📅 เลือกเดือน", ["-- แสดงทั้งหมด --"], disabled=True)


    # --- Filter แผนก
    with col3:
        departments = ["-- แสดงทั้งหมด --"] + sorted(df_filtered["แผนก"].dropna().unique())
        selected_dept = st.selectbox("🏢 เลือกแผนก", departments)
        if selected_dept != "-- แสดงทั้งหมด --":
            df_filtered = df_filtered[df_filtered["แผนก"] == selected_dept]

    # --- คำนวณประเภทการลา (ต้องเพิ่มคอลัมน์ "จำนวน" สำหรับการคำนวณยอดรวม)
    
    def leave_days(x):
        """นับวันลา: 0.5 สำหรับครึ่งวัน, 1 สำหรับวันเต็ม/ครั้ง"""
        if "ครึ่งวัน" in str(x):
            return 0.5
        # รวมข้อยกเว้นทั้งหมดที่นับเป็น 1 วัน/ครั้ง
        valid_full_day_exceptions = ["ลาป่วย", "ลากิจ", "ขาด", "สาย", "ลาพักผ่อน", "ลาคลอด"] 
        if str(x) in valid_full_day_exceptions:
            return 1
        return 0 # ค่าที่ไม่ใช่การลา/ขาด/สาย

    # ******************************************************************************
    # เพิ่มการคำนวณจำนวนวัน/ครั้งลาลงใน df_filtered เพื่อใช้ในส่วนของ Expander
    df_filtered["จำนวน"] = df_filtered["ข้อยกเว้น"].apply(leave_days)
    # ******************************************************************************

    df_filtered["ลาป่วย/ลากิจ"] = df_filtered["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ลาป่วย", "ลากิจ", "ลาป่วยครึ่งวัน", "ลากิจครึ่งวัน"] else 0
    )
    df_filtered["ขาด"] = df_filtered["ข้อยกเว้น"].apply(
        lambda x: leave_days(x) if str(x) in ["ขาด", "ขาดครึ่งวัน"] else 0
    )
    df_filtered["สาย"] = df_filtered["ข้อยกเว้น"].apply(lambda x: 1 if str(x) == "สาย" else 0)
    df_filtered["ลาพักผ่อน"] = df_filtered["ข้อยกเว้น"].apply(lambda x: 1 if str(x) == "ลาพักผ่อน" else 0)

    leave_types = ["ลาป่วย/ลากิจ", "ขาด", "สาย", "ลาพักผ่อน"]
    summary = df_filtered.groupby(["ชื่อ-สกุล", "แผนก"])[leave_types].sum().reset_index()

    st.title("📊 แดชบอร์ดการลา / ขาด / สาย")

    # --- ตัวกรองพนักงาน ---
    if 'selected_employee' not in st.session_state:
        st.session_state.selected_employee = '-- แสดงทั้งหมด --'
    all_names = ["-- แสดงทั้งหมด --"] + sorted(summary["ชื่อ-สกุล"].unique())
    selected_employee = st.selectbox("🔍 ค้นหาชื่อพนักงาน", all_names, key='selected_employee')

    colors = {
        "ลาป่วย/ลากิจ": "#FFC300",
        "ขาด": "#C70039",
        "สาย": "#FF5733",
        "ลาพักผ่อน": "#33C4FF"
    }

    tabs = st.tabs(leave_types)
    for t, leave in zip(tabs, leave_types):
        with t:
            st.subheader(f"🏆 จัดอันดับ {leave}")

            # --- กรองตามพนักงาน
            if selected_employee != "-- แสดงทั้งหมด --":
                summary_filtered = summary[summary["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
                # ใช้ df_filtered เดิม ซึ่งมีคอลัมน์ "จำนวน" แล้ว
                person_data_full = df_filtered[df_filtered["ชื่อ-สกุล"] == selected_employee].reset_index(drop=True)
            else:
                summary_filtered = summary
                person_data_full = df_filtered

            st.markdown("### 📌 สรุปข้อมูล")
            # ปรับปรุง: แสดงผลให้รองรับทศนิยม 1 ตำแหน่ง
            summary_display_copy = summary_filtered.copy()
            for col in leave_types:
                # ใช้ format_value ที่สร้างขึ้น
                summary_display_copy[col] = summary_display_copy[col].apply(lambda x: format_value(x))

            st.dataframe(summary_display_copy, use_container_width=True, hide_index=True)

            # --- แสดงรายละเอียดวันลา (ส่วนที่ถูกแก้ไข) ---
            if selected_employee != "-- แสดงทั้งหมด --" and not person_data_full.empty:
                # 1. กำหนดข้อยกเว้นที่เกี่ยวข้อง
                if leave == "ลาป่วย/ลากิจ":
                    relevant_exceptions = ["ลาป่วย", "ลากิจ", "ลาป่วยครึ่งวัน", "ลากิจครึ่งวัน"]
                    unit = 'วัน'
                elif leave == "ขาด":
                    relevant_exceptions = ["ขาด", "ขาดครึ่งวัน"]
                    unit = 'วัน'
                elif leave == "สาย":
                    relevant_exceptions = ["สาย"]
                    unit = 'ครั้ง'
                else: # ลาพักผ่อน
                    relevant_exceptions = ["ลาพักผ่อน"]
                    unit = 'วัน'

                # 2. กรองข้อมูล
                dates = person_data_full.loc[
                    person_data_full["ข้อยกเว้น"].isin(relevant_exceptions),
                    ["วันที่", "เวลาเข้า", "เวลาออก", "ข้อยกเว้น", "จำนวน"] # ดึงคอลัมน์ "จำนวน" มาด้วย
                ].sort_values(by="วันที่", ascending=False)

                if not dates.empty:
                    with st.expander(f"ดูรายละเอียดวันที่ ({leave})"):
                        # 3. คำนวณยอดรวม
                        total_days = dates["จำนวน"].sum()
                        
                        for _, row in dates.iterrows():
                            # จัดรูปแบบวันที่, เวลาเข้า, เวลาออก
                            date_str = thai_date(row['วันที่']) 
                            entry_time = row['เวลาเข้า'].strftime('%H:%M')
                            exit_time = row['เวลาออก'].strftime('%H:%M')
                            exception_text = row['ข้อยกเว้น']
                            
                            # แสดงรายการ
                            label = f"• **{date_str}** &nbsp;&nbsp; **{entry_time}** - **{exit_time}** &nbsp;&nbsp;&nbsp;&nbsp; **{exception_text}**"
                            st.markdown(label, unsafe_allow_html=True)

                        # 4. แสดงยอดรวมตามที่ต้องการ
                        # ใช้ format_value เพื่อแสดง 7.5 หรือ 7.0
                        st.markdown("---")
                        st.markdown(f"**ยอดรวม:** **{format_value(total_days)}** {unit}", unsafe_allow_html=True)


            # --- ตารางอันดับ
            ranking = summary_filtered[["ชื่อ-สกุล", "แผนก", leave]].sort_values(by=leave, ascending=False).reset_index(drop=True)
            ranking.insert(0, "อันดับ", range(1, len(ranking) + 1))
            ranking_display = ranking[ranking[leave] > 0]
            
            # ปรับปรุง: แสดงผลให้รองรับทศนิยม 1 ตำแหน่ง
            ranking_display_copy = ranking_display.copy()
            ranking_display_copy[leave] = ranking_display_copy[leave].apply(lambda x: format_value(x))

            st.markdown("### 🏅 ตารางอันดับ")
            st.dataframe(ranking_display_copy.reset_index(drop=True), use_container_width=True, hide_index=True)

            # --- กราฟ
            if not ranking_display.empty:
                chart_data = ranking_display if selected_employee != "-- แสดงทั้งหมด --" else ranking_display.head(20)
                chart_title = f"ข้อมูล {leave}" if selected_employee != "-- แสดงทั้งหมด --" else f"20 อันดับแรกของพนักงานที่ '{leave}'"

                chart = (
                    alt.Chart(chart_data)
                    .mark_bar(cornerRadius=5, color=colors.get(leave, "#C70039"))
                    .encode(
                        x=alt.X(leave + ":Q", title=f"จำนวน ({'วัน' if leave != 'สาย' else 'ครั้ง'})"),
                        y=alt.Y("ชื่อ-สกุล:N", sort="-x", title="ชื่อ-สกุล"),
                        tooltip=["อันดับ", "ชื่อ-สกุล", "แผนก", leave],
                    )
                    .properties(title=chart_title)
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info(f"ไม่พบข้อมูล '{leave}' ในช่วงเวลาที่เลือก")
else:
    st.info("กรุณาตรวจสอบว่ามีไฟล์ attendances.xlsx อยู่ในโฟลเดอร์เดียวกับโปรแกรม")