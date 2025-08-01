import pandas as pd
from datetime import datetime, timedelta
import win32evtlog, win32con
import tkinter as tk
from tkinter import ttk, messagebox

class EventLogViewer:
    def __init__(self, root):
        self.root = root

        # Set global font for Tkinter widgets
        root.option_add("*Font", ("Fira Code", 10))
        
        # Configure ttk widget fonts
        style = ttk.Style()
        style.configure("TLabel", font=("Fira Code", 10))
        style.configure("TButton", font=("Fira Code", 10))
        style.configure("TEntry", font=("Fira Code", 10))
        style.configure("TCombobox", font=("Fira Code", 10))
        style.configure("Treeview", font=("Fira Code", 10))
        style.configure("Treeview.Heading", font=("Fira Code", 10))
        style.configure("TLabelframe", font=("Fira Code", 10))
        style.configure("TLabelframe.Label", font=("Fira Code", 10))
        
        root.title("PC Log Viewer")
        root.geometry("600x900")
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.date_frame = ttk.LabelFrame(self.main_frame, text="Select Month", padding="10")
        self.date_frame.pack(fill=tk.X, pady=5)
        ttk.Label(self.date_frame, text="Year:").grid(row=0, column=0, padx=5)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.year_entry = ttk.Entry(self.date_frame, textvariable=self.year_var, width=5)
        self.year_entry.grid(row=0, column=1, padx=5)
        ttk.Label(self.date_frame, text="Month:").grid(row=0, column=2, padx=5)
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        self.month_combobox = ttk.Combobox(self.date_frame, textvariable=self.month_var,
                                           values=[str(i) for i in range(1, 13)], width=3)
        self.month_combobox.grid(row=0, column=3, padx=5)
        self.fetch_btn = ttk.Button(self.date_frame, text="Fetch Logs", command=self.fetch_logs)
        self.fetch_btn.grid(row=0, column=4, padx=10)

        self.results_frame = ttk.LabelFrame(self.main_frame, text="Daily Summary", padding="10")
        self.results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.tree = ttk.Treeview(self.results_frame,
                                 columns=('Date', 'First Startup', 'Last Shutdown', 'Work Hours'),
                                 show='headings', selectmode='browse')
        for col in ('Date', 'First Startup', 'Last Shutdown', 'Work Hours'):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=50, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(self.tree, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_var = tk.StringVar() 
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, pady=5)
        self.export_btn = ttk.Button(self.main_frame, text="Export to CSV", command=self.export_to_csv, state=tk.DISABLED)
        self.export_btn.pack(pady=5)
        self.summary_df = pd.DataFrame()

    def fetch_logs(self):
        try:
            year, month = int(self.year_var.get()), int(self.month_var.get())
            if not 1 <= month <= 12:
                raise ValueError("Month must be 1-12")
            self.status_var.set(f"Fetching logs for {year}-{month:02d}...")
            self.root.update()
            events_df = self.get_system_events_for_month(year, month)
            if events_df.empty:
                messagebox.showinfo("No Data", "No events found for specified month.")
                self.status_var.set("Ready")
                self.export_btn.config(state=tk.DISABLED)
                return
            self.summary_df = self.create_daily_summary(events_df)
            self.display_results(self.summary_df)
            self.status_var.set(f"Found {len(self.summary_df)} days with events for {year}-{month:02d}")
            self.export_btn.config(state=tk.NORMAL)
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
            self.status_var.set("Ready")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Ready")
    
    def display_results(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for _, row in df.iterrows():
            date_str = row['Date'].strftime('%Y-%m-%d (%a)')
            start = row['First Startup'].strftime('%H:%M:%S') if pd.notna(row['First Startup']) else "N/A"
            shutdown = row['Last Shutdown'].strftime('%H:%M:%S') if pd.notna(row['Last Shutdown']) else "N/A"
            if pd.notna(row['Work Hours']):
                work = row['Work Hours'] - timedelta(hours=1)
                work = work if work.total_seconds() >= 0 else timedelta(0)
                total = int(work.total_seconds())
                h, rem = divmod(total, 3600)
                m, s = divmod(rem, 60)
                work_str = f"{h:02d}:{m:02d}:{s:02d}"
            else:
                work_str = "N/A"
            self.tree.insert('', tk.END, values=(date_str, start, shutdown, work_str))
    
    def export_to_csv(self):
        if self.summary_df.empty:
            return
        base = f"pc_events_{self.year_var.get()}_{self.month_var.get()}"
        def fmt_time(t): return t.strftime('%H:%M:%S') if pd.notna(t) else "N/A"
        def fmt_date(d): return d.strftime('%Y-%m-%d (%a)') if pd.notna(d) else "N/A"
        def fmt_work(td):
            if pd.notna(td):
                adj = td - timedelta(hours=1)
                if adj.total_seconds() < 0: adj = timedelta(0)
                total = int(adj.total_seconds())
                h, rem = divmod(total, 3600)
                m, s = divmod(rem, 60)
                return f"{h:02d}:{m:02d}:{s:02d}"
            return "N/A"
        def round_td(td):
            if pd.isna(td): return td
            mins = td.total_seconds() / 60
            base = (mins // 30) * 30
            return timedelta(minutes=(base if (mins - base) < 20 else base + 30))
        def td_to_decimal(td):
            if pd.isna(td): return None
            adj = td - timedelta(hours=1)
            if adj.total_seconds() < 0: adj = timedelta(0)
            dec = adj.total_seconds() / 3600
            whole = int(dec)
            mins = (dec - whole) * 60
            return whole if mins < 25 else round(dec * 2) / 2

        df_ex = self.summary_df.copy()
        df_ex['Date'] = df_ex['Date'].apply(fmt_date)
        df_ex['First Startup'] = df_ex['First Startup'].apply(fmt_time)
        df_ex['Last Shutdown'] = df_ex['Last Shutdown'].apply(fmt_time)
        df_ex['Work Hours'] = df_ex['Work Hours'].apply(fmt_work)
        df_ex['Work Hours (Rounded)'] = self.summary_df['Work Hours'].apply(lambda td: fmt_work(round_td(td)) if pd.notna(td) else "N/A")
        df_ex['Work Hours (Decimal)'] = self.summary_df['Work Hours'].apply(lambda td: td_to_decimal(td) if pd.notna(td) else None)
        df_ex = df_ex[['Date', 'First Startup', 'Last Shutdown', 'Work Hours', 'Work Hours (Rounded)', 'Work Hours (Decimal)']]
        df_ex.to_csv(f"{base}.csv", index=False)
        messagebox.showinfo("Export Complete", f"File exported: {base}.csv")
    
    @staticmethod
    def get_system_events_for_month(year, month):
        first_day = datetime(year, month, 1)
        last_day = datetime(year+1, 1, 1) - timedelta(days=1) if month == 12 else datetime(year, month+1, 1) - timedelta(days=1)
        start, end = first_day, last_day + timedelta(days=1)
        hand = win32evtlog.OpenEventLog(None, "System")
        win32evtlog.ReadEventLog(hand, win32con.EVENTLOG_BACKWARDS_READ | win32con.EVENTLOG_SEQUENTIAL_READ, 0)
        events = []
        flags = win32con.EVENTLOG_BACKWARDS_READ | win32con.EVENTLOG_SEQUENTIAL_READ
        try:
            while True:
                recs = win32evtlog.ReadEventLog(hand, flags, 0)
                if not recs:
                    break
                for r in recs:
                    evt = r.EventID & 0xFFFF
                    if evt in (6005, 6006):
                        t = r.TimeGenerated
                        if start <= t <= end:
                            events.append({'Date': t.date(), 'Time': t.time(), 'EventID': evt,
                                           'EventType': 'Startup' if evt == 6005 else 'Shutdown'})
        finally:
            win32evtlog.CloseEventLog(hand)
        df = pd.DataFrame(events)
        return df.sort_values(by=['Date', 'Time']) if not df.empty else df
    
    @staticmethod
    def create_daily_summary(df):
        if df.empty:
            return pd.DataFrame(columns=['Date', 'First Startup', 'Last Shutdown', 'Work Hours'])
        st = df[df['EventID'] == 6005].groupby('Date').first()
        sh = df[df['EventID'] == 6006].groupby('Date').last()
        summary = pd.DataFrame({'Date': st.index, 'First Startup': st['Time'], 'Last Shutdown': sh['Time']})
        all_dates = pd.date_range(min(df['Date']), max(df['Date']), freq='D').date
        summary = summary.set_index('Date').reindex(all_dates).reset_index().rename(columns={'index': 'Date'})
        summary['Work Hours'] = summary.apply(lambda row: datetime.combine(row['Date'], row['Last Shutdown']) - datetime.combine(row['Date'], row['First Startup']) if pd.notna(row['First Startup']) and pd.notna(row['Last Shutdown']) else pd.NaT, axis=1)
        return summary

if __name__ == "__main__":
    root = tk.Tk()
    app = EventLogViewer(root)
    root.mainloop()
