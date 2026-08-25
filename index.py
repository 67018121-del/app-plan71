import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
import requests
import base64

EXCEL_FILE = 'รายการแผนคอม 71.xlsx'

GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", os.environ.get("GITHUB_TOKEN", ""))
REPO_NAME = st.secrets.get("REPO_NAME", "67018121-del/app-plan71")

st.set_page_config(page_title="ระบบแผนคอมพิวเตอร์ 71", layout="wide")

# --- ฟังก์ชันบันทึกกลับไปยัง GitHub อัตโนมัติ ---
def save_to_github(file_bytes, commit_message):
    if not GITHUB_TOKEN:
        st.error("⚠️ ไม่พบ GITHUB_TOKEN กรุณาตั้งค่า Secrets ใน Streamlit Cloud")
        return False

    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{EXCEL_FILE}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    # 1. ดึง sha ของไฟล์เดิม
    res = requests.get(url, headers=headers)
    sha = res.json().get("sha", "") if res.status_code == 200 else ""
    
    # 2. ส่งข้อมูลอัปเดตไฟล์กลับไป
    content_b64 = base64.b64encode(file_bytes).decode("utf-8")
    payload = {
        "message": commit_message,
        "content": content_b64
    }
    if sha:
        payload["sha"] = sha
    
    put_res = requests.put(url, json=payload, headers=headers)
    if put_res.status_code not in [200, 201]:
        st.error(f"GitHub API Error: {put_res.status_code} - {put_res.json().get('message', '')}")
        return False
    return True

# --- ฟังก์ชันอ่านข้อมูลจากไฟล์ Excel ---
def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            df_main = pd.read_excel(EXCEL_FILE, sheet_name=0, engine='openpyxl').fillna('')
            try:
                df_log = pd.read_excel(EXCEL_FILE, sheet_name='ประวัติการแก้ไข', engine='openpyxl').fillna('')
            except:
                df_log = pd.DataFrame()
            return df_main, df_log
        except Exception as e:
            st.error(f"Error loading excel file: {e}")
    return pd.DataFrame(), pd.DataFrame()

# โหลดข้อมูลเข้าสู่ Session State
if 'df_main' not in st.session_state or 'df_log' not in st.session_state:
    df_m, df_l = load_data()
    st.session_state.df_main = df_m
    st.session_state.df_log = df_l

df_main = st.session_state.df_main
df_log = st.session_state.df_log

st.title("💻 ระบบค้นหา แก้ไข และยืนยันข้อมูลแผนคอมพิวเตอร์ 71")

if not df_main.empty:
    # ==========================================
    # 1. ข้อมูลผู้ดำเนินการ
    # ==========================================
    st.subheader("👤 ข้อมูลผู้ทำรายการ")
    col1, col2, col3, col4 = st.columns(4)
    u_name = col1.text_input("ชื่อ-นามสกุล:")
    u_id = col2.text_input("รหัสพนักงาน:")
    u_pos = col3.text_input("ตำแหน่ง:")
    u_dept = col4.text_input("หน่วยงานผู้ทำรายการ:")

    # ==========================================
    # 2. ตัวกรองและค้นหา
    # ==========================================
    st.subheader("🔍 ตัวกรองและค้นหา")
    col_f1, col_f2, col_f3 = st.columns(3)
    
    depts = ["-- ทั้งหมด --"] + sorted([str(x).strip() for x in df_main['หน่วยงาน'].unique() if str(x).strip()])
    types = ["-- ทั้งหมด --"] + sorted([str(x).strip() for x in df_main['รายการ/ประเภท'].unique() if str(x).strip()])
    
    selected_dept = col_f1.selectbox("หน่วยงาน:", depts)
    selected_type = col_f2.selectbox("รายการ/ประเภท:", types)
    search_txt = col_f3.text_input("ค้นหาคำ:")

    # กรองข้อมูล
    filtered_df = df_main.copy()
    if selected_dept != "-- ทั้งหมด --":
        filtered_df = filtered_df[filtered_df['หน่วยงาน'].astype(str).str.strip() == selected_dept]
    if selected_type != "-- ทั้งหมด --":
        filtered_df = filtered_df[filtered_df['รายการ/ประเภท'].astype(str).str.strip() == selected_type]
    if search_txt:
        filtered_df = filtered_df[filtered_df.astype(str).apply(lambda x: x.str.contains(search_txt, case=False)).any(axis=1)]

    # ==========================================
    # 3. แบบสอบถามความต้องการยืนยัน/แก้ไขข้อมูล
    # ==========================================
    st.subheader("📋 เลือกสถานะรายการ")
    
    if not filtered_df.empty:
        options = {f"ลำดับ {row.get('ลำดับ', idx + 1)}: [{row.get('หน่วยงาน', '')}] {row.get('รายการ/ประเภท', '')}": idx for idx, row in filtered_df.iterrows()}
        selected_option = st.selectbox("เลือกรายการที่จะดำเนินการ:", list(options.keys()))
        
        selected_idx = options[selected_option]
        current_row = df_main.loc[selected_idx]
        
        # แบบสอบถามความต้องการ
        action_type = st.radio(
            "หน่วยงานต้องการแก้ไขจำนวน หรือยืนยันข้อมูลตามเดิม?",
            ["✅ ยืนยันข้อมูลตามเดิม (ไม่แก้ไข)", "✏️ ต้องการแก้ไขจำนวนขอทดแทน"],
            horizontal=True
        )

        # ----------------------------------------------------
        # กรณีที่ 1: กด "ยืนยันข้อมูลตามเดิม"
        # ----------------------------------------------------
        if action_type == "✅ ยืนยันข้อมูลตามเดิม (ไม่แก้ไข)":
            st.info(f"📌 รายการปัจจุบัน: ขอใหม่ **{current_row.get('ขอใหม่', 0)}** เครื่อง | ขอทดแทน **{current_row.get('ขอทดแทน', 0)}** เครื่อง | รวม **{current_row.get('รวม', 0)}** เครื่อง (งบประมาณ **{current_row.get('จำนวนเงิน', 0):,.2f}** บาท)")
            
            if st.button("✅ ยืนยันข้อมูลถูกต้อง", type="primary"):
                if not u_name or not u_id or not u_pos or not u_dept:
                    st.warning("⚠️ กรุณากรอกข้อมูลผู้ทำรายการให้ครบถ้วนก่อนกดบันทึก")
                else:
                    curr_replace = int(current_row.get('ขอทดแทน', 0) or 0)
                    curr_amount = float(current_row.get('จำนวนเงิน', 0) or 0)
                    unit_price = float(current_row.get('ราคาต่อหน่วย', 0) or 0)

                    # สร้าง Log สำหรับการยืนยัน
                    log_entry = {
                        'เวลาที่แก้ไข': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'ลำดับ': current_row.get('ลำดับ', selected_idx + 1),
                        'หน่วยงาน': current_row.get('หน่วยงาน', ''),
                        'รายการ/ประเภท': current_row.get('รายการ/ประเภท', ''),
                        'ขอทดแทน (เดิม)': curr_replace,
                        'ขอทดแทน (ใหม่)': curr_replace,
                        'ผลต่างจำนวน': 0,
                        'ราคาต่อหน่วย': unit_price,
                        'จำนวนเงินเดิม': curr_amount,
                        'จำนวนเงินใหม่': curr_amount,
                        'ผลต่างงบประมาณ': 0,
                        'ชื่อ-นามสกุล ผู้แก้ไข': u_name,
                        'รหัสพนักงาน': u_id,
                        'ตำแหน่ง': u_pos,
                        'หน่วยงานผู้แก้ไข': u_dept,
                        'หมายเหตุ': 'ยืนยันข้อมูลตามเดิม'
                    }

                    new_log_df = pd.DataFrame([log_entry])
                    st.session_state.df_log = pd.concat([st.session_state.df_log, new_log_df], ignore_index=True)

                    # บันทึกเข้า GitHub
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        st.session_state.df_main.to_excel(writer, sheet_name='ข้อมูลแผนคอมพิวเตอร์', index=False)
                        st.session_state.df_log.to_excel(writer, sheet_name='ประวัติการแก้ไข', index=False)
                    excel_bytes = output.getvalue()

                    success = save_to_github(excel_bytes, f"Confirmed by {u_name} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    
                    if success:
                        st.success(f"✅ บันทึกการ 'ยืนยันข้อมูลตามเดิม' สำเร็จเรียบร้อยแล้วโดย {u_name}")
                        st.rerun()

        # ----------------------------------------------------
        # กรณีที่ 2: กด "ต้องการแก้ไขจำนวนขอทดแทน"
        # ----------------------------------------------------
        else:
            val_replace_new = st.number_input("ขอทดแทน (ระบุจำนวนใหม่):", value=int(current_row.get('ขอทดแทน', 0)), step=1, min_value=0, max_value=999)

            if st.button("💾 บันทึกการแก้ไขข้อมูล", type="primary"):
                if not u_name or not u_id or not u_pos or not u_dept:
                    st.warning("⚠️ กรุณากรอกข้อมูลผู้ทำรายการให้ครบถ้วนก่อนกดบันทึก")
                else:
                    old_replace = int(current_row.get('ขอทดแทน', 0) or 0)
                    old_amount = float(current_row.get('จำนวนเงิน', 0) or 0)
                    unit_price = float(current_row.get('ราคาต่อหน่วย', 0) or 0)
                    val_new = int(current_row.get('ขอใหม่', 0) or 0)

                    if old_replace == val_replace_new:
                        st.info("ℹ️ จำนวนขอทดแทนไม่ได้เปลี่ยนแปลง")
                    else:
                        val_total_new = val_new + val_replace_new
                        new_amount = val_total_new * unit_price

                        # 1. อัปเดตข้อมูล Sheet หลัก
                        st.session_state.df_main.loc[selected_idx, 'ขอทดแทน'] = val_replace_new
                        st.session_state.df_main.loc[selected_idx, 'รวม'] = val_total_new
                        st.session_state.df_main.loc[selected_idx, 'จำนวนเงิน'] = new_amount

                        # 2. สร้าง Log บันทึกประวัติการแก้ไข
                        log_entry = {
                            'เวลาที่แก้ไข': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'ลำดับ': current_row.get('ลำดับ', selected_idx + 1),
                            'หน่วยงาน': current_row.get('หน่วยงาน', ''),
                            'รายการ/ประเภท': current_row.get('รายการ/ประเภท', ''),
                            'ขอทดแทน (เดิม)': old_replace,
                            'ขอทดแทน (ใหม่)': val_replace_new,
                            'ผลต่างจำนวน': val_replace_new - old_replace,
                            'ราคาต่อหน่วย': unit_price,
                            'จำนวนเงินเดิม': old_amount,
                            'จำนวนเงินใหม่': new_amount,
                            'ผลต่างงบประมาณ': new_amount - old_amount,
                            'ชื่อ-นามสกุล ผู้แก้ไข': u_name,
                            'รหัสพนักงาน': u_id,
                            'ตำแหน่ง': u_pos,
                            'หน่วยงานผู้แก้ไข': u_dept,
                            'หมายเหตุ': 'มีการแก้ไขจำนวน'
                        }

                        new_log_df = pd.DataFrame([log_entry])
                        st.session_state.df_log = pd.concat([st.session_state.df_log, new_log_df], ignore_index=True)

                        # 3. บันทึกเข้า GitHub
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            st.session_state.df_main.to_excel(writer, sheet_name='ข้อมูลแผนคอมพิวเตอร์', index=False)
                            st.session_state.df_log.to_excel(writer, sheet_name='ประวัติการแก้ไข', index=False)
                        excel_bytes = output.getvalue()

                        success = save_to_github(excel_bytes, f"Updated by {u_name} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                        
                        if success:
                            st.success(f"✅ บันทึกการแก้ไขข้อมูลสำเร็จเรียบร้อยแล้วโดย {u_name}")
                            st.rerun()

    # ==========================================
    # 4. สรุปยอดเงิน, ตารางแสดงผล & ปุ่มดาวน์โหลด
    # ==========================================
    total_budget = filtered_df['จำนวนเงิน'].astype(float).sum()
    st.markdown(f"### 💰 สรุปรวมงบประมาณ: **{total_budget:,.2f}** บาท")

    st.dataframe(filtered_df, use_container_width=True)

    output_download = io.BytesIO()
    with pd.ExcelWriter(output_download, engine='openpyxl') as writer:
        st.session_state.df_main.to_excel(writer, sheet_name='ข้อมูลแผนคอมพิวเตอร์', index=False)
        st.session_state.df_log.to_excel(writer, sheet_name='ประวัติการแก้ไข', index=False)
    
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Excel ล่าสุด (รวมประวัติการยืนยัน/แก้ไข)",
        data=output_download.getvalue(),
        file_name="รายการแผนคอม 71_อัปเดต.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.error("ไม่พบข้อมูลในไฟล์ Excel หรือไฟล์เสียหาย โปรดตรวจสอบไฟล์ 'รายการแผนคอม 71.xlsx'")
