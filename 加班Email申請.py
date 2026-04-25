import tkinter as tk
from tkinter import messagebox, ttk
import win32com.client as win32
from datetime import datetime
import time
import os
import threading
import pythoncom
import re
import shutil

class OvertimeEmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("辦公自動化：加班申請助手") # 改為較正式的名稱
        self.root.geometry("480x550") 
        self.root.configure(bg="#f5f5f5")

        # --- 初始化設定 (可根據需求修改) ---
        today = datetime.now()
        self.roc_year = today.year - 1911 # 台灣民國年格式
        self.current_month = today.month
        
        self.week_options = ["一", "二", "三", "四", "五", "六", "日"]
        current_week_day = self.week_options[today.weekday()]

        self.date_var = tk.StringVar(value=f"民國{self.roc_year}年{today.month}月{today.day}日")
        self.week_var = tk.StringVar(value=current_week_day) 
        self.time_var = tk.StringVar(value="17:30~20:30")
        self.user_name_var = tk.StringVar(value="YourName") # TODO: 預設使用者姓名
        
        # --- 敏感資訊佔位符 ---
        self.receiver_email = "manager@example.com" # TODO: 設定主管信箱
        self.cc_email = "your_backup@example.com"   # TODO: 設定副本信箱
        self.is_monitoring = False 
        self.send_time_anchor = None 

        # --- UI 介面 ---
        tk.Label(root, text="✉️ 加班申請自動化系統", font=('Microsoft JhengHei', 16, 'bold'), bg="#f5f5f5").pack(pady=20)
        main_frame = tk.Frame(root, bg="#f5f5f5")
        main_frame.pack(padx=30, fill="x")

        # 欄位配置
        self._create_label_entry(main_frame, "【申請人姓名】:", self.user_name_var)
        self._create_label_entry(main_frame, "【加班日期】:", self.date_var)
        
        tk.Label(main_frame, text="【加班星期】:", bg="#f5f5f5", font=('Microsoft JhengHei', 10, 'bold')).pack(anchor="w")
        self.week_menu = ttk.Combobox(main_frame, textvariable=self.week_var, values=self.week_options, state="readonly", font=('Arial', 11), width=38)
        self.week_menu.pack(pady=(0, 10))

        self._create_label_entry(main_frame, "【加班時間】:", self.time_var)

        tk.Label(main_frame, text="【加班事由】:", bg="#f5f5f5", font=('Microsoft JhengHei', 10, 'bold')).pack(anchor="w")
        self.reason_text = tk.Text(main_frame, font=('Microsoft JhengHei', 10), width=40, height=5)
        self.reason_text.pack(pady=(0, 10))
        self.reason_text.insert("1.0", "處理專案自動化系統開發與測試") # 通用事由

        self.status_label = tk.Label(root, text="系統狀態: 待命", bg="#f5f5f5", fg="gray", font=('Microsoft JhengHei', 9))
        self.status_label.pack(pady=5)

        self.send_btn = tk.Button(root, text="發送申請並啟動回信監控", command=self.send_and_start_monitor,
                                 bg="#007bff", fg="white", font=('Microsoft JhengHei', 12, 'bold'),
                                 height=2, width=35, cursor="hand2")
        self.send_btn.pack(pady=15)

    def _create_label_entry(self, parent, label_text, var):
        tk.Label(parent, text=label_text, bg="#f5f5f5", font=('Microsoft JhengHei', 10, 'bold')).pack(anchor="w")
        tk.Entry(parent, textvariable=var, font=('Arial', 11), width=40).pack(pady=(0, 10))

    def get_monthly_details_fast(self):
        """從 Outlook 已傳送郵件中統計本月加班次數"""
        try:
            outlook = win32.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            sent_folder = ns.GetDefaultFolder(5) # 5 = Sent Mail
            this_month_start = datetime(datetime.now().year, datetime.now().month, 1).strftime("%m/%d/%Y %H:%M %p")
            filter_str = f"[SentOn] >= '{this_month_start}'"
            filtered_items = sent_folder.Items.Restrict(filter_str)
            
            count = 0
            date_list = set()
            for msg in filtered_items:
                try:
                    subj = str(msg.Subject)
                    if "加班申請" in subj:
                        match = re.search(r'(\d+)/(\d+)/(\d+)', subj)
                        if match:
                            m_val = int(match.group(2))
                            if m_val == self.current_month:
                                count += 1
                                date_list.add(f"{str(m_val).zfill(2)}/{str(match.group(3)).zfill(2)}")
                except: continue
            return count, list(date_list)
        except: return 0, []

    def get_funny_comment(self, count):
        if count <= 2: return "🌱 初級努力家：工作之餘也要記得休息喔！"
        elif count <= 5: return "🔥 辦公室精英：你是團隊中不可或缺的力量！"
        else: return "🚀 核心貢獻者：感謝你的付出，但健康才是最重要的財富！"

    def send_and_start_monitor(self):
        self.send_btn.config(state="disabled")
        self.status_label.config(text="系統狀態: ⚡ 正在處理數據...", fg="#d9534f")
        self.root.update()

        try:
            hist_count, hist_dates = self.get_monthly_details_fast()
            self.send_time_anchor = datetime.now()
            
            # 解析輸入日期
            raw_input_date = self.date_var.get()
            nums = re.findall(r'\d+', raw_input_date)
            formatted_date = f"{nums[0]}/{nums[1]}/{nums[2]}" if len(nums) >= 3 else "DateError"
            
            target_tag = f"{formatted_date}加班申請_{self.user_name_var.get().strip()}"
            
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To, mail.CC = self.receiver_email, self.cc_email
            mail.Subject = target_tag 
            
            # 保留預設簽名檔並插入內容
            mail.Display()
            signature = mail.HTMLBody
            body_content = (f"<p>主管您好，<br><br>"
                            f"【加班日期】：{self.date_var.get()} (星期{self.week_var.get()}) <br>"
                            f"【加班時間】：{self.time_var.get()}<br>"
                            f"【加班事由】：{self.reason_text.get('1.0', 'end-1c')}<br><br>"
                            f"謝謝!</p>")
            mail.HTMLBody = body_content + signature
            mail.Send()
            
            # 統計更新
            total_count = hist_count + 1
            comment = self.get_funny_comment(total_count)
            messagebox.showinfo("成功", f"申請信已送出！\n本月累積紀錄：{total_count} 次\n\n{comment}")
            
            if not self.is_monitoring:
                self.is_monitoring = True
                self.status_label.config(text="系統狀態: 🔍 監控回信中...", fg="#0056b3")
                threading.Thread(target=self.background_monitor, args=(target_tag.upper(), self.send_time_anchor), daemon=True).start()
        except Exception as e:
            messagebox.showerror("錯誤", f"發送失敗：{e}")
            self.send_btn.config(state="normal")

    def background_monitor(self, target_tag_upper, send_time_anchor):
        pythoncom.CoInitialize()
        try:
            # 尋找桌面路徑 (通用邏輯)
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            
            temp_dir = r"C:\Temp_Mail_Export"
            if not os.path.exists(temp_dir): os.makedirs(temp_dir)

            while self.is_monitoring:
                outlook_app = win32.Dispatch("Outlook.Application")
                ns = outlook_app.GetNamespace("MAPI")
                inbox = ns.GetDefaultFolder(6) # 6 = Inbox
                
                after_send_str = send_time_anchor.strftime("%m/%d/%Y %H:%M %p")
                items = inbox.Items.Restrict(f"[ReceivedTime] >= '{after_send_str}'")
                items.Sort("[ReceivedTime]", True) 

                for i in range(1, min(10, items.Count + 1)):
                    msg = items.Item(i)
                    try:
                        subj_raw = str(msg.Subject).upper()
                        if "RE:" in subj_raw and target_tag_upper in subj_raw:
                            msg_time = msg.ReceivedTime.replace(tzinfo=None)
                            if msg_time > send_time_anchor:
                                # 存檔邏輯
                                clean_name = re.sub(r'[\\/*?:"<>|]', "_", str(msg.Subject))
                                final_name = f"{clean_name}_{msg_time.strftime('%H%M%S')}.msg"
                                save_path = os.path.join(desktop, final_name)
                                msg.SaveAs(save_path, 3) # 3 = olMsg
                                self.root.after(0, self.finish_monitor)
                                return
                    except: continue
                time.sleep(15) 
        finally: pythoncom.CoUninitialize()

    def finish_monitor(self):
        self.is_monitoring = False
        self.status_label.config(text="系統狀態: ✅ 已收到回覆", fg="#28a745")
        self.send_btn.config(state="normal")
        messagebox.showinfo("通知", "主管已回信！郵件已自動存至桌面。")

if __name__ == "__main__":
    root = tk.Tk()
    app = OvertimeEmailApp(root)
    root.mainloop()
