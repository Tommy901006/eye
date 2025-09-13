import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
from scipy.stats import entropy
import smtplib
from email.message import EmailMessage

class CrossEntropyGUI:
    def __init__(self, master):
        self.master = master
        master.title("Batch Cross Entropy Calculator")
        master.geometry("650x500")

        # 啟動時詢問收件者 Email（可跳過）
        self.recipient_email = simpledialog.askstring("收件人 Email", "請輸入收件者 Email（可跳過）:")

        self.selected_cols = []
        self.file_columns = []
        self.combo_cols = []

        self.build_interface()

    def build_interface(self):
        frame_path = ttk.LabelFrame(self.master, text="Folder and Column Selection", padding=10)
        frame_path.pack(fill='x', padx=10, pady=5)

        ttk.Button(frame_path, text="Select Folder", command=self.select_folder).pack(anchor='w')
        self.lbl_folder = ttk.Label(frame_path, text="")
        self.lbl_folder.pack(anchor='w', pady=5)

        frame_cols = ttk.Frame(frame_path)
        frame_cols.pack()
        for i in range(2):
            ttk.Label(frame_cols, text=f"Column {i+1}:").grid(row=i, column=0, sticky='e')
            cb = ttk.Combobox(frame_cols, width=30, state="readonly")
            cb.grid(row=i, column=1, padx=5, pady=2)
            self.combo_cols.append(cb)

        frame_log = ttk.LabelFrame(self.master, text="Execution Log", padding=10)
        frame_log.pack(fill='both', expand=True, padx=10, pady=5)

        self.log = scrolledtext.ScrolledText(frame_log, height=12)
        self.log.pack(fill='both', expand=True)

        ttk.Button(self.master, text="Start Batch Processing", command=self.start_processing).pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.lbl_folder.config(text=folder)
            files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.csv'))]
            if files:
                df = pd.read_excel(os.path.join(folder, files[0])) if files[0].endswith('xlsx') else pd.read_csv(os.path.join(folder, files[0]))
                self.file_columns = df.columns.tolist()
                for cb in self.combo_cols:
                    cb['values'] = [''] + self.file_columns
                    cb.set('')
            else:
                messagebox.showwarning("No Files", "No Excel or CSV files found in the folder.")

    def log_message(self, msg):
        self.log.insert(tk.END, msg + '\n')
        self.log.yview(tk.END)

    def calculate_cross_entropy(self, col_x, col_y):
        if col_x.nunique() < 2:
            return np.nan, f"Column 1 only has one unique value: {col_x.unique()}"
        if col_y.nunique() < 2:
            return np.nan, f"Column 2 only has one unique value: {col_y.unique()}"

        value_counts_x = col_x.value_counts()
        value_counts_y = col_y.value_counts()

        prob_x = value_counts_x / len(col_x)
        prob_y = value_counts_y / len(col_y)

        prob_x, prob_y = prob_x.align(prob_y, fill_value=0)

        nonzero_indices = prob_y > 0
        if nonzero_indices.sum() == 0:
            return np.nan, "No overlapping nonzero values between columns."

        prob_x_filtered = prob_x[nonzero_indices]
        prob_y_filtered = prob_y[nonzero_indices]

        if prob_x_filtered.sum() == 0 or prob_y_filtered.sum() == 0:
            return np.nan, "Zero sum after filtering probabilities."

        prob_x_filtered = prob_x_filtered / prob_x_filtered.sum()
        prob_y_filtered = prob_y_filtered / prob_y_filtered.sum()

        epsilon = 1e-10
        pred_prob = np.clip(prob_y_filtered.values, epsilon, 1. - epsilon)
        true_prob = prob_x_filtered.values

        cross_entropy_value = -np.sum(true_prob * np.log2(pred_prob))
        return cross_entropy_value, None

    def start_processing(self):
        folder = self.lbl_folder.cget("text")
        selected_cols = [cb.get() for cb in self.combo_cols if cb.get()]

        if not folder or len(selected_cols) != 2:
            messagebox.showerror("Missing Information", "Please select a folder and exactly two columns.")
            return

        results = []
        files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.csv'))]
        for file in files:
            try:
                path = os.path.join(folder, file)
                df = pd.read_excel(path) if file.endswith('xlsx') else pd.read_csv(path)

                if selected_cols[0] not in df.columns or selected_cols[1] not in df.columns:
                    self.log_message(f"{file}: Columns not found.")
                    continue

                col_x = df[selected_cols[0]].dropna()
                col_y = df[selected_cols[1]].dropna()

                min_len = min(len(col_x), len(col_y))
                col_x = col_x.iloc[:min_len]
                col_y = col_y.iloc[:min_len]

                col_x = pd.cut(col_x, bins=10, labels=False)
                col_y = pd.cut(col_y, bins=10, labels=False)

                cross_entropy_value, reason = self.calculate_cross_entropy(col_x, col_y)

                results.append({
                    "File Name": file,
                    "Cross_Entropy": cross_entropy_value,
                    "Reason": reason if reason else "OK",
                    "Unique_Col1": col_x.nunique(),
                    "Unique_Col2": col_y.nunique()
                })

                if reason:
                    self.log_message(f"Processed {file}: Cross Entropy = nan ({reason})")
                else:
                    self.log_message(f"Processed {file}: Cross Entropy = {cross_entropy_value:.4f}")

            except Exception as e:
                self.log_message(f"Error processing {file}: {e}")

        if results:
            out_path = os.path.join(folder, "Cross_Entropy_Results.xlsx")
            df_results = pd.DataFrame(results)
            df_results.to_excel(out_path, index=False)
            messagebox.showinfo("Completed", f"Results saved to {out_path}")

            # ✅ 如果有輸入 Email，自動寄送結果檔案
            if self.recipient_email:
                self.send_email(self.recipient_email, out_path)

            try:
                os.startfile(folder)
            except Exception as e:
                self.log_message(f"Failed to open folder: {e}")

    def send_email(self, to_email, file_path):
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = "4A930099@stust.edu.tw"  # 🔒請改成你自己的 Gmail
        sender_password = "trwvoyligttxqcjy"     # 🔒App 密碼（非登入密碼）

        try:
            msg = EmailMessage()
            msg['Subject'] = 'Cross Entropy 分析結果'
            msg['From'] = sender_email
            msg['To'] = to_email
            msg.set_content("您好，附件為 Cross Entropy 統計分析結果 Excel 檔案。")

            with open(file_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(file_path)
                msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)

            self.log_message("📧 Email 已成功寄出至 " + to_email)
        except Exception as e:
            self.log_message("❌ Email 發送失敗：" + str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = CrossEntropyGUI(root)
    root.mainloop()
