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
    
    res = requests.get(url, headers=headers)
    sha = res.json().get("sha", "") if res.status_code == 200 else ""
    
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
            
            # บันทึกค่า Max ลิมิตตั้งต้น
            if 'ขอทดแทน_MAX' not in df_main.columns:
                df_main['ขอทดแทน_MAX'] = pd.to_numeric(df_main['ขอทดแทน'], errors='coerce').fillna(0).astype(int)
                
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

st.title("💻 ระบบตรวจสอบ แก้ไข และยืนยันข้อมูลแผนคอมพิวเตอร์ 71")

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
    # 2. ตัวกรองเลือกหน่วยงาน
    # ==========================================
    st.subheader("🔍 เลือกหน่วยงาน")
    depts = ["-- ทั้งหมด --"] + sorted([str(x).strip() for x in df_main['หน่วยงาน'].unique() if str(x).strip()])
    selected_dept = st.selectbox("เลือกหน่วยงานที่ต้องการตรวจสอบ:", depts)

    # กรองข้อมูล
    if selected_dept != "-- ทั้งหมด --":
        filter_mask = df_main['หน่วยงาน'].astype(str).str.strip() == selected_dept
    else:
        filter_mask = pd.Series([True] * len(df_main))

    filtered_indices = df_main[filter_mask].index.tolist()

    # Calculate Total Budget
    total_budget = df_main.loc[filtered_indices, 'จำนวนเงิน'].astype(float).sum()
    st.markdown(f"### 💰 สรุปรวมงบประมาณหน่วยงานที่เลือก: **{total_budget:,.2f}** บาท")

    st.markdown("---")
    st.subheader("📋 รายการข้อมูลแผนคอมพิวเตอร์")

    # Header ของรายการ
    h_c1, h_c2, h_c3, h_c4, h_c5, h_c6, h_c7 = st.columns([1, 4, 2, 2, 2, 3, 3])
    h_c1.markdown("**ลำดับ**")
    h_c2.markdown("**รายการ/ประเภท**")
    h_c3.markdown("**ขอใหม่**")
    h_c4.markdown("**ขอทดแทน**")
    h_c5.markdown("**รวม**")
    h_c6.markdown("**งบประมาณ (บาท)**")
    h_c7.markdown("**สถานะ / การดำเนินการ**")

    # ==========================================
    # 3. แสดงรายการทีละแถวพร้อมปุ่มดำเนินการข้างตาราง
    # ==========================================
    has_changes = False

    for idx in filtered_indices:
        row = st.session_state.df_main.loc[idx]
        
        c1, c2, c3, c4, c5, c6, c7 = st.columns([1, 4, 2, 2, 2, 3, 3])
        
        c1.write(f"{row.get('ลำดับ', idx + 1)}")
        c2.write(f"{row.get('รายการ/ประเภท', '')}")
        c3.write(f"{row.get('ขอใหม่', 0)}")
        
        # แสดงขอทดแทนปัจจุบัน
        curr_replace = int(row.get('ขอทดแทน', 0))
        max_limit = int(row.get('ขอทดแทน_MAX', curr_replace))
        c4.write(f"**{curr_replace}** (Max: {max_limit})")
        
        c5.write(f"{row.get('รวม', 0)}")
        c6.write(f"{float(row.get('จำนวนเงิน', 0)):,.2f}")

        # ฝั่งขวา: สวิตช์เลือกว่า ยืนยัน หรือ แก้ไข
        action = c7.radio(
            f"action_{idx}",
            ["ยืนยันข้อมูลเดิม", "แก้ไข"],
            key=f"radio_{idx}",
            horizontal=True,
            label_visibility="collapsed"
        )

        # หากเลือก "แก้ไข" ให้แสดงปุ่ม + / - ข้างตาราง
        if action == "แก้ไข":
            btn_col1, btn_col2, btn_col3 = c7.columns([1, 2, 1])
            
            # ปุ่มลบ (-)
            if btn_col1.button("➖", key=f"dec_{idx}", disabled=(curr_replace <= 0)):
                new_replace = curr_replace - 1
                new_total = int(row.get('ขอใหม่', 0)) + new_replace
                unit_price = float(row.get('ราคาต่อหน่วย', 0))
                
                st.session_state.df_main.loc[idx, 'ขอทดแทน'] = new_replace
                st.session_state.df_main.loc[idx, 'รวม'] = new_total
                st.session_state.df_main.loc[idx, 'จำนวนเงิน'] = new_total * unit_price
                st.rerun()

            btn_col2.markdown(f"<h5 style='text-align: center; margin: 0;'>{curr_replace}</h5>", unsafe_allow_html=True)

            # ปุ่มเพิ่ม (+)
            if btn_col3.button("➕", key=f"inc_{idx}", disabled=(curr_replace >= max_limit)):
                new_replace = curr_replace + 1
                new_total = int(row.get('ขอใหม่', 0)) + new_replace
                unit_price = float(row.get('ราคาต่อหน่วย', 0))
                
                st.session_state.df_main.loc[idx, 'ขอทดแทน'] = new_replace
                st.session_state.df_main.loc[idx, 'รวม'] = new_total
                st.session_state.df_main.loc[idx, 'จำนวนเงิน'] = new_total * unit_price
                st.rerun()

        st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)

    # ==========================================
    # 4. ปุ่มบันทึกข้อมูลและดาวน์โหลด
    # ==========================================
    st.markdown("---")
    col_save, col_dl = st.columns([1, 1])

    with col_save:
        if st.button("💾 บันทึกการทำรายการลง GitHub", type="primary", use_container_width=True):
            if not u_name or not u_id or not u_pos or not u_dept:
                st.warning("⚠️ กรุณากรอกข้อมูลผู้ทำรายการให้ครบถ้วนด้านบนก่อนกดบันทึก")
            else:
                # สร้าง Log สรุปการรายการ
                new_logs = []
                for idx in filtered_indices:
                    row = st.session_state.df_main.loc[idx]
                    action_val = st.session_state.get(f"radio_{idx}", "ยืนยันข้อมูลเดิม")
                    
                    orig_replace = int(row.get('ขอทดแทน_MAX', 0))
                    curr_rep = int(row.get('ขอทดแทน', 0))
                    unit_p = float(row.get('ราคาต่อหน่วย', 0))
                    
                    log_entry = {
                        'เวลาที่แก้ไข': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'ลำดับ': row.get('ลำดับ', idx + 1),
                        'หน่วยงาน': row.get('หน่วยงาน', ''),
                        'รายการ/ประเภท': row.get('รายการ/ประเภท', ''),
                        'ขอทดแทน (เดิม)': orig_replace,
                        'ขอทดแทน (ใหม่)': curr_rep,
                        'ผลต่างจำนวน': curr_rep - orig_replace,
                        'ราคาต่อหน่วย': unit_p,
                        'จำนวนเงินใหม่': float(row.get('จำนวนเงิน', 0)),
                        'ชื่อ-นามสกุล ผู้แก้ไข': u_name,
                        'รหัสพนักงาน': u_id,
                        'ตำแหน่ง': u_pos,
                        'หน่วยงานผู้แก้ไข': u_dept,
                        'หมายเหตุ': action_val
                    }
                    new_logs.append(log_entry)

                new_log_df = pd.DataFrame(new_logs)
                st.session_state.df_log = pd.concat([st.session_state.df_log, new_log_df], ignore_index=True)

                # แปลงไฟล์เซฟลง GitHub
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_save = st.session_state.df_main.drop(columns=['ขอทดแทน_MAX'], errors='ignore')
                    df_save.to_excel(writer, sheet_name='ข้อมูลแผนคอมพิวเตอร์', index=False)
                    st.session_state.df_log.to_excel(writer, sheet_name='ประวัติการแก้ไข', index=False)
                excel_bytes = output.getvalue()

                success = save_to_github(excel_bytes, f"Updated/Confirmed by {u_name} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                if success:
                    st.success("✅ บันทึกข้อมูลลง GitHub เรียบร้อยแล้ว!")
                    st.rerun()

    with col_dl:
        output_download = io.BytesIO()
        with pd.ExcelWriter(output_download, engine='openpyxl') as writer:
            df_save = st.session_state.df_main.drop(columns=['ขอทดแทน_MAX'], errors='ignore')
            df_save.to_excel(writer, sheet_name='ข้อมูลแผนคอมพิวเตอร์', index=False)
            st.session_state.df_log.to_excel(writer, sheet_name='ประวัติการแก้ไข', index=False)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel ล่าสุด",
            data=output_download.getvalue(),
            file_name="รายการแผนคอม 71_อัปเดต.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

else:
    st.error("ไม่พบข้อมูลในไฟล์ Excel หรือไฟล์เสียหาย โปรดตรวจสอบไฟล์ 'รายการแผนคอม 71.xlsx'")
