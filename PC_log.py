import pandas as pd
from datetime import datetime, timedelta
import win32evtlog
import win32con
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar

class EventLogViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("PC Startup/Shutdown Log Viewer")
        self.root.geometry("800x900")
        
        # Create main container
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Date selection
        self.date_frame = ttk.LabelFrame(self.main_frame, text="Select Month", padding="10")
        self.date_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.date_frame, text="Year:").grid(row=0, column=0, padx=5)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.year_entry = ttk.Entry(self.date_frame, textvariable=self.year_var, width=5)
        self.year_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(self.date_frame, text="Month:").grid(row=0, column=2, padx=5)
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        self.month_combobox = ttk.Combobox(self.date_frame, textvariable=self.month_var, 
                                          values = [str(i) for i in range(1, 13)], width=3)
        self.month_combobox.grid(row=0, column=3, padx=5)
        
        self.fetch_btn = ttk.Button(self.date_frame, text="Fetch Logs", command=self.fetch_logs)
        self.fetch_btn.grid(row=0, column=4, padx=10)
        
        # Calendar selection alternative
        self.calendar_btn = ttk.Button(self.date_frame, text="Select from Calendar", 
                                      command=self.show_calendar)
        self.calendar_btn.grid(row=0, column=5, padx=10)
        
        # Results display
        self.results_frame = ttk.LabelFrame(self.main_frame, text="Daily Summary", padding="10")
        self.results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Treeview for results, now including a "Work Hours" column
        self.tree = ttk.Treeview(self.results_frame, columns=('Date', 'First Startup', 'Last Shutdown', 'Work Hours'), 
                                show='headings', selectmode='browse')
        
        self.tree.heading('Date', text='Date')
        self.tree.heading('First Startup', text='First Startup')
        self.tree.heading('Last Shutdown', text='Last Shutdown')
        self.tree.heading('Work Hours', text='Work Hours')
        
        self.tree.column('Date', width=100, anchor=tk.W)
        self.tree.column('First Startup', width=100, anchor=tk.CENTER)
        self.tree.column('Last Shutdown', width=100, anchor=tk.CENTER)
        self.tree.column('Work Hours', width=100, anchor=tk.CENTER)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.tree, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, 
                                  relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, pady=5)
        
        # Export button
        self.export_btn = ttk.Button(self.main_frame, text="Export to CSV", 
                                    command=self.export_to_csv, state=tk.DISABLED)
        self.export_btn.pack(pady=5)
        
        # Initialize variables
        self.summary_df = pd.DataFrame()
    
    def show_calendar(self):
        """Show calendar popup for date selection with EXE compatibility"""
        Calendar = None
        try:
            # Try to use tkcalendar first
            import tkcalendar
            Calendar = tkcalendar.Calendar
            use_native = False
        except ImportError:
            # Fallback to simpler date entry if tkcalendar fails
            use_native = True
        
        if not use_native and Calendar is not None:
            # tkcalendar implementation
            def on_date_select():
                try:
                    selected_date = cal.selection_get()
                    if selected_date is not None:
                        self.year_var.set(str(selected_date.year))
                        self.month_var.set(str(selected_date.month))
                    top.destroy()
                except:
                    top.destroy()
            
            top = tk.Toplevel(self.root)
            top.title("Select Date")
            top.geometry("300x250")
            top.transient(self.root)
            top.grab_set()
            
            try:
                year = int(self.year_var.get())
                month = int(self.month_var.get())
            except:
                year = datetime.now().year
                month = datetime.now().month
            
            cal = Calendar(top, selectmode='day', year=year, month=month, locale='en_US')
            cal.pack(padx=10, pady=10)
            
            button_frame = ttk.Frame(top)
            button_frame.pack(pady=5, anchor='center')
            ttk.Button(button_frame, text="Select", command=on_date_select).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="Cancel", command=top.destroy).pack(side=tk.LEFT, padx=5)
            self.root.wait_window(top)
        else:
            # Fallback implementation
            self.fallback_date_selection()

    def fallback_date_selection(self):
        """Fallback date selection when tkcalendar is not available"""
        top = tk.Toplevel(self.root)
        top.title("Select Month/Year")
        top.geometry("300x150")
        
        frame = ttk.Frame(top, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Year:").grid(row=0, column=0, padx=5, pady=5)
        year_entry = ttk.Entry(frame, width=6)
        year_entry.grid(row=0, column=1, padx=5, pady=5)
        year_entry.insert(0, self.year_var.get())
        
        ttk.Label(frame, text="Month:").grid(row=1, column=0, padx=5, pady=5)
        month_combobox = ttk.Combobox(frame, values=[str(i) for i in range(1, 13)], width=4)
        month_combobox.grid(row=1, column=1, padx=5, pady=5)
        month_combobox.set(self.month_var.get())
        
        def set_date():
            try:
                self.year_var.set(str(int(year_entry.get())))
                self.month_var.set(str(int(month_combobox.get())))
                top.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers")
        
        ttk.Button(frame, text="OK", command=set_date).grid(row=2, column=0, columnspan=2, pady=10)
        top.transient(self.root)
        top.grab_set()
        self.root.wait_window(top)
    
    def fetch_logs(self):
        """Fetch event logs based on selected month"""
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
            
            if not (1 <= month <= 12):
                raise ValueError("Month must be between 1 and 12")
            
            self.status_var.set(f"Fetching logs for {year}-{month:02d}...")
            self.root.update()
            
            events_df = self.get_system_events_for_month(year, month)
            
            if events_df.empty:
                messagebox.showinfo("No Data", "No startup/shutdown events found for the specified month.")
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
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set("Ready")
    
    def display_results(self, df):
        """Display results in the Treeview"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insert new data
        for _, row in df.iterrows():
            date_str = row['Date'].strftime('%Y-%m-%d (%A)')
            startup_str = row['First Startup'].strftime('%H:%M:%S') if pd.notna(row['First Startup']) else "N/A"
            shutdown_str = row['Last Shutdown'].strftime('%H:%M:%S') if pd.notna(row['Last Shutdown']) else "N/A"
            if pd.notna(row['Work Hours']):
                # Subtract 1 hour for lunch break
                adjusted_work_duration = row['Work Hours'] - timedelta(hours=1)
                # Ensure non-negative duration
                if adjusted_work_duration.total_seconds() < 0:
                    adjusted_work_duration = timedelta(seconds=0)
                total_seconds = int(adjusted_work_duration.total_seconds())
                hours, remainder = divmod(total_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                work_hours_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            else:
                work_hours_str = "N/A"
            self.tree.insert('', tk.END, values=(date_str, startup_str, shutdown_str, work_hours_str))
    
    def export_to_csv(self):
        """Export current results to a CSV file with 6 columns:
        1. Original Date, First Startup, Last Shutdown, and Work Hours (with lunch break deducted)
        2. Work Hours (Rounded to the nearest 30 minutes as a time)
        3. Work Hours (Rounded to a decimal number)
        """
        if self.summary_df.empty:
            return

        base_path = f"pc_events_{self.year_var.get()}_{self.month_var.get()}"

        # Helper functions for formatting
        def format_time(time_val):
            return time_val.strftime('%H:%M:%S') if pd.notna(time_val) else "N/A"

        def format_date(d):
            return d.strftime('%Y-%m-%d (%A)') if pd.notna(d) else "N/A"

        def format_work_hours(td):
            # Deduct one hour for lunch break; ensure non-negative duration.
            if pd.notna(td):
                adjusted = td - timedelta(hours=1)
                if adjusted.total_seconds() < 0:
                    adjusted = timedelta(seconds=0)
                total_seconds = int(adjusted.total_seconds())
                hours, remainder = divmod(total_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            else:
                return "N/A"

        # Rounding logic for work hours (rounded to nearest 30 minutes)
        def round_timedelta(td):
            if pd.isna(td):
                return td
            total_minutes = td.total_seconds() / 60
            lower = (total_minutes // 30) * 30
            if (total_minutes - lower) < 20:
                rounded_minutes = lower
            else:
                rounded_minutes = lower + 30
            return timedelta(minutes=rounded_minutes)

        # Rounding logic to convert work hours to a decimal number
        def timedelta_to_decimal(td):
            if pd.isna(td):
                    return None
            adjusted = td - timedelta(hours=1)
            if adjusted.total_seconds() < 0:
                    adjusted = timedelta(seconds=0)
            hours_decimal = adjusted.total_seconds() / 3600
            whole_hours = int(hours_decimal)
            minutes = (hours_decimal - whole_hours) * 60
            if minutes < 25:
                return whole_hours
            else:
                return round(hours_decimal * 2) / 2

        # Create a copy of the summary for export
        df_export = self.summary_df.copy()

        # Format the original columns for export
        df_export['Date'] = df_export['Date'].apply(format_date)
        df_export['First Startup'] = df_export['First Startup'].apply(format_time)
        df_export['Last Shutdown'] = df_export['Last Shutdown'].apply(format_time)
        df_export['Work Hours'] = df_export['Work Hours'].apply(format_work_hours)

        # Create new columns for the rounded values
        df_export['Work Hours (Rounded)'] = self.summary_df['Work Hours'].apply(
            lambda td: format_work_hours(round_timedelta(td)) if pd.notna(td) else "N/A"
        )
        df_export['Work Hours (Decimal)'] = self.summary_df['Work Hours'].apply(
            lambda td: timedelta_to_decimal(td) if pd.notna(td) else None
        )

        # Arrange columns in the desired order
        df_export = df_export[['Date', 'First Startup', 'Last Shutdown',
                               'Work Hours', 'Work Hours (Rounded)', 'Work Hours (Decimal)']]

        df_export.to_csv(f"{base_path}.csv", index=False)
        messagebox.showinfo("Export Complete", f"File exported: {base_path}.csv")
    
    @staticmethod
    def get_system_events_for_month(year, month):
        """Retrieve system startup (6005) and shutdown (6006) events for a specific month."""
        # Calculate date range for the month
        first_day = datetime(year, month, 1)
        if month == 12:
            last_day = datetime(year+1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(year, month+1, 1) - timedelta(days=1)
        
        # Convert to Windows FILETIME format
        start_time = first_day
        end_time = last_day + timedelta(days=1)  # Include the entire last day
        
        # Open the System event log
        hand = win32evtlog.OpenEventLog(None, "System")
        win32evtlog.ReadEventLog(hand, win32con.EVENTLOG_BACKWARDS_READ|win32con.EVENTLOG_SEQUENTIAL_READ, 0)
        
        events = []
        flags = win32con.EVENTLOG_BACKWARDS_READ | win32con.EVENTLOG_SEQUENTIAL_READ
        
        try:
            while True:
                # Read events in chunks
                records = win32evtlog.ReadEventLog(hand, flags, 0)
                if not records:
                    break
                    
                for record in records:
                    event_id = record.EventID & 0xFFFF  # Get the base event ID
                    
                    # Check if it's one of our target events
                    if event_id in (6005, 6006):
                        event_time = record.TimeGenerated
                        
                        # Filter by date range
                        if start_time <= event_time <= end_time:
                            events.append({
                                'Date': event_time.date(),
                                'Time': event_time.time(),
                                'EventID': event_id,
                                'EventType': 'Startup' if event_id == 6005 else 'Shutdown'
                            })
        
        finally:
            win32evtlog.CloseEventLog(hand)
        
        # Create DataFrame and sort by time
        df = pd.DataFrame(events)
        if not df.empty:
            df = df.sort_values(by=['Date', 'Time'])
        
        return df
    
    @staticmethod
    def create_daily_summary(df):
        """Create a daily summary table with first startup, last shutdown, and work hours."""
        if df.empty:
            return pd.DataFrame(columns=['Date', 'First Startup', 'Last Shutdown', 'Work Hours'])
        
        # Get first startup (6005) of each day
        startups = df[df['EventID'] == 6005].groupby('Date').first()
        # Get last shutdown (6006) of each day
        shutdowns = df[df['EventID'] == 6006].groupby('Date').last()
        
        # Merge the two
        summary = pd.DataFrame({
            'Date': startups.index,
            'First Startup': startups['Time'],
            'Last Shutdown': shutdowns['Time']
        })
        
        # Ensure all dates in the month are represented
        min_date = df['Date'].min()
        max_date = df['Date'].max()
        all_dates = pd.date_range(start=min_date, end=max_date, freq='D').date
        
        summary = summary.set_index('Date').reindex(all_dates).reset_index()
        summary = summary.rename(columns={'index': 'Date'})
        
        # Calculate Work Hours if both First Startup and Last Shutdown are present
        def calc_work_hours(row):
            if pd.notna(row['First Startup']) and pd.notna(row['Last Shutdown']):
                start_dt = datetime.combine(row['Date'], row['First Startup'])
                shutdown_dt = datetime.combine(row['Date'], row['Last Shutdown'])
                return shutdown_dt - start_dt
            else:
                return pd.NaT
        
        summary['Work Hours'] = summary.apply(calc_work_hours, axis=1)
        
        return summary

if __name__ == "__main__":
    root = tk.Tk()
    app = EventLogViewer(root)
    root.mainloop()
