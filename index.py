import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import os
from datetime import datetime

EXCEL_FILE = 'รายการแผนคอม 71.xlsx'

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ระบบค้นหา แก้ไข และจำแนกข้อมูลแผนคอมพิวเตอร์ 71")
        self.geometry("1180x850")
        self.configure(bg="#F4F6F9")

        # --- อ่านไฟล์ Excel อัตโนมัติ ---
        self.raw_data = self.load_data()

        # --- ตั้งค่า Theme & Font ที่คมชัด ---
        self.setup_styles()

        # ==========================================
        # 1. ส่วนบน: ข้อมูลผู้แก้ไข & ตัวกรอง
        # ==========================================
        user_frame = ttk.LabelFrame(self, text=" 👤 ข้อมูลผู้ทำการแก้ไข ", padding=(12, 8))
        user_frame.pack(fill="x", padx=15, pady=(10, 4))

        ttk.Label(user_frame, text="ชื่อ-นามสกุล:").grid(row=0, column=0, padx=(5, 2), pady=3, sticky="e")
        self.ent_user_name = ttk.Entry(user_frame, width=24)
        self.ent_user_name.grid(row=0, column=1, padx=(0, 20), pady=3, sticky="w")

        ttk.Label(user_frame, text="รหัสพนักงาน:").grid(row=0, column=2, padx=(5, 2), pady=3, sticky="e")
        self.ent_user_id = ttk.Entry(user_frame, width=18)
        self.ent_user_id.grid(row=0, column=3, padx=(0, 20), pady=3, sticky="w")

        ttk.Label(user_frame, text="ตำแหน่ง:").grid(row=1, column=0, padx=(5, 2), pady=3, sticky="e")
        self.ent_user_pos = ttk.Entry(user_frame, width=24)
        self.ent_user_pos.grid(row=1, column=1, padx=(0, 20), pady=3, sticky="w")

        ttk.Label(user_frame, text="หน่วยงานผู้แก้ไข:").grid(row=1, column=2, padx=(5, 2), pady=3, sticky="e")
        self.ent_user_dept = ttk.Entry(user_frame, width=24)
        self.ent_user_dept.grid(row=1, column=3, padx=(0, 20), pady=3, sticky="w")

        # --- ตัวกรองและค้นหา ---
        ctrl_frame = ttk.LabelFrame(self, text=" 🔍 ตัวกรองและค้นหา ", padding=(12, 8))
        ctrl_frame.pack(fill="x", padx=15, pady=4)

        ttk.Label(ctrl_frame, text="หน่วยงาน:").grid(row=0, column=0, padx=(5, 2), pady=2)
        self.dept_combo = ttk.Combobox(ctrl_frame, state="readonly", width=20)
        self.dept_combo.grid(row=0, column=1, padx=(0, 15), pady=2)
        self.dept_combo.bind("<<ComboboxSelected>>", self.apply_filter)

        ttk.Label(ctrl_frame, text="รายการ/ประเภท:").grid(row=0, column=2, padx=(5, 2), pady=2)
        self.type_combo = ttk.Combobox(ctrl_frame, state="readonly", width=20)
        self.type_combo.grid(row=0, column=3, padx=(0, 15), pady=2)
        self.type_combo.bind("<<ComboboxSelected>>", self.apply_filter)

        ttk.Label(ctrl_frame, text="ค้นหา:").grid(row=0, column=4, padx=(5, 2), pady=2)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.apply_filter())
        ttk.Entry(ctrl_frame, textvariable=self.search_var, width=18).grid(row=0, column=5, padx=5, pady=2)

        # ==========================================
        # 2. ส่วนแก้ไขข้อมูล & สรุปยอดเงิน (ย้ายกลับมาตำแหน่งเดิม)
        # ==========================================
        # --- แก้ไขจำนวนขอทดแทน ---
        edit_frame = ttk.LabelFrame(self, text=" ✏️ แก้ไขจำนวนเครื่องขอทดแทน ", padding=(12, 8))
        edit_frame.pack(fill="x", padx=15, pady=4)

        ttk.Label(edit_frame, text="รายการที่เลือก:").grid(row=0, column=0, padx=5, sticky="w")
        self.lbl_selected_item = ttk.Label(
            edit_frame, 
            text="- โปรดคลิกเลือกรายการในตารางด้านล่างเพื่อแก้ไข -", 
            font=("Segoe UI", 10, "bold"), 
            foreground="#1976D2"
        )
        self.lbl_selected_item.grid(row=0, column=1, columnspan=3, padx=5, sticky="w")

        ttk.Label(edit_frame, text="ขอทดแทน:").grid(row=1, column=0, padx=5, pady=6, sticky="w")
        self.spin_replace = ttk.Spinbox(edit_frame, from_=0, to=999, width=10)
        self.spin_replace.grid(row=1, column=1, padx=5, pady=6, sticky="w")

        btn_update = ttk.Button(edit_frame, text="บันทึกการแก้ไข", command=self.update_item, style="Primary.TButton")
        btn_update.grid(row=1, column=2, padx=10, pady=6)

        btn_save_all = ttk.Button(edit_frame, text="💾 บันทึกลง Excel (รวม Sheet ประวัติ)", command=self.save_all_to_single_excel, style="Success.TButton")
        btn_save_all.grid(row=1, column=3, padx=10, pady=6)

        # --- สรุปยอดเงิน ---
        sum_frame = tk.Frame(self, bg="#E8F5E9", bd=1, relief="solid")
        sum_frame.pack(fill="x", padx=15, pady=6)
        
        self.lbl_summary = tk.Label(
            sum_frame, 
            text="สรุปรวมงบประมาณ: 0.00 บาท", 
            font=("Segoe UI", 12, "bold"), 
            fg="#2E7D32", 
            bg="#E8F5E9"
        )
        self.lbl_summary.pack(side="left", padx=15, pady=8)

        # ==========================================
        # 3. ส่วนด้านล่าง: ตารางแสดงผลข้อมูล
        # ==========================================
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=15, pady=(4, 15))

        columns = ("no", "type", "dept", "replace", "total", "unit_price", "amount")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        # กำหนดให้จัดกึ่งกลาง (center) สำหรับตัวเลขทุกคอลัมน์
        headers = [
            ("no", "ลำดับ", 65, "center"),
            ("type", "รายการ/ประเภท", 290, "w"),
            ("dept", "หน่วยงาน", 160, "center"),
            ("replace", "ขอทดแทน", 80, "center"),
            ("total", "รวม", 80, "center"),
            ("unit_price", "ราคาต่อหน่วย", 130, "center"),
            ("amount", "จำนวนเงิน (บาท)", 150, "center")
        ]

        for col, text, width, align in headers:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=align)

        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd", background="#F8F9FA")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<ButtonRelease-1>", self.on_select_item)
        self.selected_row_index = None

        self.setup_comboboxes()
        self.apply_filter()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        font_main = ("Segoe UI", 10)
        font_bold = ("Segoe UI", 10, "bold")
        bg_main = "#F4F6F9"

        style.configure(".", font=font_main, background=bg_main)
        style.configure("TLabelframe", background=bg_main, bordercolor="#D1D5DB", relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", font=font_bold, foreground="#1F2937", background=bg_main)
        style.configure("TLabel", background=bg_main, foreground="#374151")
        
        style.configure("TEntry", padding=4, bordercolor="#D1D5DB")
        style.configure("TCombobox", padding=3)
        style.configure("TSpinbox", padding=3)
        
        style.configure("Treeview", 
                        font=("Segoe UI", 10), 
                        rowheight=28, 
                        background="#FFFFFF", 
                        fieldbackground="#FFFFFF",
                        bordercolor="#E5E7EB")
        style.configure("Treeview.Heading", 
                        font=("Segoe UI", 10, "bold"), 
                        background="#E5E7EB", 
                        foreground="#1F2937", 
                        padding=5,
                        relief="flat")
        style.map("Treeview", background=[("selected", "#E0F2FE")], foreground=[("selected", "#0369A1")])
        
        style.configure("TButton", font=font_bold, padding=(10, 5), borderwidth=0)
        style.configure("Primary.TButton", background="#1976D2", foreground="white")
        style.map("Primary.TButton", background=[("active", "#1565C0")])

        style.configure("Success.TButton", background="#2E7D32", foreground="white")
        style.map("Success.TButton", background=[("active", "#1B5E20")])

    def load_data(self):
        if os.path.exists(EXCEL_FILE):
            try:
                df = pd.read_excel(EXCEL_FILE, sheet_name=0)
                df = df.fillna('')
                try:
                    df_log = pd.read_excel(EXCEL_FILE, sheet_name='ประวัติการแก้ไข')
                    self.modified_history = df_log.to_dict(orient='records')
                except:
                    self.modified_history = []
                return df.to_dict(orient='records')
            except Exception as e:
                print(f"Error loading excel file: {e}")
                return []
        self.modified_history = []
        return []

    def setup_comboboxes(self):
        depts = sorted(list({str(item.get('หน่วยงาน', '')).strip() for item in self.raw_data if item.get('หน่วยงาน')}))
        types = sorted(list({str(item.get('รายการ/ประเภท', '')).strip() for item in self.raw_data if item.get('รายการ/ประเภท')}))

        self.dept_combo['values'] = ["-- ทั้งหมด --"] + depts
        self.dept_combo.current(0)

        self.type_combo['values'] = ["-- ทั้งหมด --"] + types
        self.type_combo.current(0)

    def apply_filter(self, *args):
        selected_dept = self.dept_combo.get()
        selected_type = self.type_combo.get()
        search_txt = self.search_var.get().lower().strip()

        for item in self.tree.get_children():
            self.tree.delete(item)

        total_budget = 0.0
        row_count = 0

        for idx, row in enumerate(self.raw_data):
            dept_match = (selected_dept == "-- ทั้งหมด --") or (str(row.get('หน่วยงาน', '')).strip() == selected_dept)
            type_match = (selected_type == "-- ทั้งหมด --") or (str(row.get('รายการ/ประเภท', '')).strip() == selected_type)
            search_match = any(search_txt in str(v).lower() for v in row.values()) if search_txt else True

            if dept_match and type_match and search_match:
                try:
                    amount = float(row.get('จำนวนเงิน', 0) or 0)
                except:
                    amount = 0.0

                try:
                    unit_price = float(row.get('ราคาต่อหน่วย', 0) or 0)
                except:
                    unit_price = 0.0

                total_budget += amount
                tag = "odd" if row_count % 2 == 1 else "even"

                self.tree.insert("", "end", iid=str(idx), values=(
                    row.get('ลำดับ', idx + 1),
                    row.get('รายการ/ประเภท', ''),
                    row.get('หน่วยงาน', ''),
                    row.get('ขอทดแทน', 0),
                    row.get('รวม', 0),
                    f"{unit_price:,.2f}",
                    f"{amount:,.2f}"
                ), tags=(tag,))
                
                row_count += 1

        self.lbl_summary.config(text=f"สรุปรวมงบประมาณ: {total_budget:,.2f} บาท")

    def on_select_item(self, event):
        selected_iid = self.tree.focus()
        if not selected_iid:
            return
        
        self.selected_row_index = int(selected_iid)
        item_data = self.raw_data[self.selected_row_index]

        title = f"[{item_data.get('หน่วยงาน', '')}] {item_data.get('รายการ/ประเภท', '')}"
        self.lbl_selected_item.config(text=title)

        self.spin_replace.delete(0, "end")
        self.spin_replace.insert(0, str(item_data.get('ขอทดแทน', 0)))

    def update_item(self):
        if self.selected_row_index is None:
            messagebox.showwarning("คำเตือน", "กรุณาคลิกเลือกรายการในตารางก่อนแก้ไขครับ")
            return

        u_name = self.ent_user_name.get().strip()
        u_id = self.ent_user_id.get().strip()
        u_pos = self.ent_user_pos.get().strip()
        u_dept = self.ent_user_dept.get().strip()

        if not u_name or not u_id or not u_pos or not u_dept:
            messagebox.showwarning("คำเตือน", "กรุณากรอกข้อมูลผู้แก้ไขให้ครบถ้วน (ชื่อ-นามสกุล, รหัสพนักงาน, ตำแหน่ง, หน่วยงานผู้แก้ไข)")
            return

        try:
            val_replace_new = int(self.spin_replace.get() or 0)
        except ValueError:
            messagebox.showerror("ผิดพลาด", "กรุณากรอกตัวเลขจำนวนเครื่องขอทดแทนให้ถูกต้อง")
            return

        row = self.raw_data[self.selected_row_index]
        old_replace = int(row.get('ขอทดแทน', 0) or 0)
        old_amount = float(row.get('จำนวนเงิน', 0) or 0)
        unit_price = float(row.get('ราคาต่อหน่วย', 0) or 0)
        val_new = int(row.get('ขอใหม่', 0) or 0)

        if old_replace == val_replace_new:
            messagebox.showinfo("แจ้งเตือน", "จำนวนขอทดแทนไม่ได้เปลี่ยนแปลง")
            return

        val_total_new = val_new + val_replace_new
        new_amount = val_total_new * unit_price

        log_entry = {
            'เวลาที่แก้ไข': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'ชื่อ-นามสกุล ผู้แก้ไข': u_name,
            'รหัสพนักงาน': u_id,
            'ตำแหน่ง': u_pos,
            'หน่วยงานผู้แก้ไข': u_dept,
            'ลำดับ': row.get('ลำดับ', self.selected_row_index + 1),
            'หน่วยงาน': row.get('หน่วยงาน', ''),
            'รายการ/ประเภท': row.get('รายการ/ประเภท', ''),
            'ขอทดแทน (เดิม)': old_replace,
            'ขอทดแทน (ใหม่)': val_replace_new,
            'ผลต่างจำนวน': val_replace_new - old_replace,
            'ราคาต่อหน่วย': unit_price,
            'จำนวนเงินเดิม': old_amount,
            'จำนวนเงินใหม่': new_amount,
            'ผลต่างงบประมาณ': new_amount - old_amount
        }
        self.modified_history.append(log_entry)

        row['ขอทดแทน'] = val_replace_new
        row['รวม'] = val_total_new
        row['จำนวนเงิน'] = new_amount

        self.apply_filter()
        messagebox.showinfo("สำเร็จ", f"อัปเดตข้อมูลเรียบร้อยโดย {u_name}! อย่าลืมกดปุ่มบันทึกลงไฟล์ Excel")

    def save_all_to_single_excel(self):
        if not self.raw_data:
            return
        try:
            df_main = pd.DataFrame(self.raw_data)
            df_log = pd.DataFrame(self.modified_history) if self.modified_history else pd.DataFrame()

            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                df_main.to_excel(writer, sheet_name='ข้อมูลแผนคอมพิวเตอร์', index=False)
                if not df_log.empty:
                    df_log.to_excel(writer, sheet_name='ประวัติการแก้ไข', index=False)

            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลและประวัติการแก้ไขรวมลงในไฟล์ '{EXCEL_FILE}' เรียบร้อยแล้ว!")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"ไม่สามารถบันทึกไฟล์ Excel ได้: {e}\n(หมายเหตุ: โปรดปิดไฟล์ Excel ก่อนกดบันทึก)")

if __name__ == "__main__":
    app = App()
    app.mainloop()