import streamlit as st
import pandas as pd
from datetime import datetime
import os

EXCEL_FILE = 'รายการแผนคอม 71.xlsx'

st.set_page_config(page_title="ระบบแผนคอมพิวเตอร์ 71", layout="wide")
st.title("💻 ระบบค้นหา แก้ไข และจำแนกข้อมูลแผนคอมพิวเตอร์ 71")

# --- อ่านข้อมูล ---
@st.cache_data(ttl=1)
def load_data():
    if os.path.exists(EXCEL_FILE):
        df_main = pd.read_excel(EXCEL_FILE, sheet_name=0).fillna('')
        try:
            df_log = pd.read_excel(EXCEL_FILE, sheet_name='ประวัติการแก้ไข').fillna('')
        except:
            df_log = pd.DataFrame()
        return df_main, df_log
    return pd.DataFrame(), pd.DataFrame()

df_main, df_log = load_data()

if not df_main.empty:
    # 1. ข้อมูลผู้แก้ไข
    with st.expander("👤 ข้อมูลผู้ทำการแก้ไข", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        u_name = col1.text_input("ชื่อ-นามสกุล")
        u_id = col2.text_input("รหัสพนักงาน")
        u_pos = col3.text_input("ตำแหน่ง")
        u_dept = col4.text_input("หน่วยงานผู้แก้ไข")

    # 2. ตัวกรองข้อมูล
    with st.expander("🔍 ตัวกรองและค้นหา", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        depts = ["-- ทั้งหมด --"] + sorted(df_main['หน่วยงาน'].astype(str).unique().tolist())
        types = ["-- ทั้งหมด --"] + sorted(df_main['รายการ/ประเภท'].astype(str).unique().tolist())
        
        sel_dept = col_f1.selectbox("หน่วยงาน", depts)
        sel_type = col_f2.selectbox("รายการ/ประเภท", types)
        search_txt = col_f3.text_input("ค้นหาคำ")

    # กรองข้อมูล
    filtered_df = df_main.copy()
    if sel_dept != "-- ทั้งหมด --":
        filtered_df = filtered_df[filtered_df['หน่วยงาน'] == sel_dept]
    if sel_type != "-- ทั้งหมด --":
        filtered_df = filtered_df[filtered_df['รายการ/ประเภท'] == sel_type]
    if search_txt:
        filtered_df = filtered_df[filtered_df.astype(str).apply(lambda x: x.str.contains(search_txt, case=False)).any(axis=1)]

    # 3. แก้ไขข้อมูล
    with st.expander("✏️ แก้ไขจำนวนเครื่องขอทดแทน", expanded=True):
        if not filtered_df.empty:
            item_options = filtered_df.apply(lambda r: f"ลำดับ {r.name + 1}: [{r['หน่วยงาน']}] {r['รายการ/ประเภท']}", axis=1)
            selected_item = st.selectbox("เลือกรายการที่จะแก้ไข", item_options)
            
            row_idx = int(selected_item.split(":")[0].replace("ลำดับ ", "")) - 1
            current_row = df_main.iloc[row_idx]
            
            new_replace = st.number_input("จำนวนขอทดแทน", value=int(current_row.get('ขอทดแทน', 0)), step=1)

            if st.button("💾 บันทึกการแก้ไขลง Excel"):
                if not (u_name and u_id and u_pos and u_dept):
                    st.error("กรุณากรอกข้อมูลผู้แก้ไขให้ครบถ้วนก่อนบันทึก")
                else:
                    # อัปเดตข้อมูล
                    old_replace = int(current_row.get('ขอทดแทน', 0))
                    old_amount = float(current_row.get('จำนวนเงิน', 0))
                    unit_price = float(current_row.get('ราคาต่อหน่วย', 0))
                    val_new = int(current_row.get('ขอใหม่', 0))

                    val_total_new = val_new + new_replace
                    new_amount = val_total_new * unit_price

                    # บันทึกประวัติ
                    new_log = {
                        'เวลาที่แก้ไข': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'ชื่อ-นามสกุล ผู้แก้ไข': u_name, 'รหัสพนักงาน': u_id,
                        'ตำแหน่ง': u_pos, 'หน่วยงานผู้แก้ไข': u_dept,
                        'ลำดับ': row_idx + 1, 'หน่วยงานรายการ': current_row['หน่วยงาน'],
                        'รายการ/ประเภท': current_row['รายการ/ประเภท'],
                        'ขอทดแทน (เดิม)': old_replace, 'ขอทดแทน (ใหม่)': new_replace,
                        'ผลต่างจำนวน': new_replace - old_replace,
                        'ราคาต่อหน่วย': unit_price, 'จำนวนเงินเดิม': old_amount,
                        'จำนวนเงินใหม่': new_amount, 'ผลต่างงบประมาณ': new_amount - old_amount
                    }

                    df_main.at[row_idx, 'ขอทดแทน'] = new_replace
                    df_main.at[row_idx, 'รวม'] = val_total_new
                    df_main.at[row_idx, 'จำนวนเงิน'] = new_amount

                    df_log = pd.concat([df_log, pd.DataFrame([new_log])], ignore_index=True)

                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                        df_main.to_excel(writer, sheet_name='ข้อมูลแผนคอมพิวเตอร์', index=False)
                        df_log.to_excel(writer, sheet_name='ประวัติการแก้ไข', index=False)

                    st.success("บันทึกข้อมูลเรียบร้อย!")
                    st.cache_data.clear()

    # 4. สรุปงบประมาณ
    total_budget = filtered_df['จำนวนเงิน'].sum()
    st.info(f"### 💰 สรุปรวมงบประมาณ: {total_budget:,.2f} บาท")

    # 5. ตารางข้อมูล
    st.dataframe(filtered_df, use_container_width=True)
else:
    st.warning("ไม่พบไฟล์ Excel หรือไม่มีข้อมูล")
