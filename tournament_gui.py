import xml.etree.ElementTree as ET
import csv
from datetime import datetime
from collections import defaultdict
import os
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

class TournamentAttendanceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Tournament Attendance Analyzer")
        self.root.geometry("600x500")
        
        self.folder_path = tk.StringVar()
        self.file_pattern = tk.StringVar(value="*.tdf")
        self.output_filename = tk.StringVar(value="tournament_attendance.csv")
        self.target_month = tk.StringVar(value="All")
        self.target_year = tk.StringVar(value="All")
        self.custom_output_path = tk.StringVar()
        self.output_location = tk.StringVar(value="Same as tournament files folder")
        
        self.setup_gui()
    
    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="Tournament Attendance Analyzer", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="5")
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.log_text = tk.Text(log_frame, height=10, width=70)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        ttk.Label(main_frame, text="Tournament Files Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.folder_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_folder).grid(row=1, column=2, pady=5)
        
        ttk.Label(main_frame, text="File Pattern:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.file_pattern, width=20).grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(main_frame, text="(e.g., *.tdf, *.xml)", font=('Arial', 8)).grid(row=2, column=2, sticky=tk.W, pady=5)
        
        date_frame = ttk.LabelFrame(main_frame, text="Date Filters", padding="10")
        date_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(date_frame, text="Month:").grid(row=0, column=0, sticky=tk.W, padx=5)
        month_combo = ttk.Combobox(date_frame, textvariable=self.target_month, width=15)
        month_combo['values'] = ('All', 'January', 'February', 'March', 'April', 'May', 'June',
                                'July', 'August', 'September', 'October', 'November', 'December')
        month_combo.grid(row=0, column=1, padx=5)
        month_combo.state(['readonly'])
        month_combo.bind('<<ComboboxSelected>>', lambda e: self.update_filename())
        
        ttk.Label(date_frame, text="Year:").grid(row=0, column=2, sticky=tk.W, padx=5)
        year_combo = ttk.Combobox(date_frame, textvariable=self.target_year, width=15)
        current_year = datetime.now().year
        year_combo['values'] = ('All',) + tuple(str(year) for year in range(current_year-5, current_year+2))
        year_combo.grid(row=0, column=3, padx=5)
        year_combo.state(['readonly'])
        year_combo.bind('<<ComboboxSelected>>', lambda e: self.update_filename())
        
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="10")
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(output_frame, text="Output Filename:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Entry(output_frame, textvariable=self.output_filename, width=30).grid(row=0, column=1, padx=5)
        ttk.Button(output_frame, text="Choose Location", command=self.choose_output_location).grid(row=0, column=2, padx=5)
        
        output_location_label = ttk.Label(output_frame, textvariable=self.output_location, 
                                        font=('Arial', 8), foreground='gray')
        output_location_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Preview Files", command=self.preview_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Generate Report", command=self.generate_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
    
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.update_filename()
            self.log(f"Selected folder: {folder}")
            
    def log(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def generate_filename(self):
        base_name = "tournament_attendance"
        
        if self.target_month.get() != "All":
            month_num = self.get_month_number(self.target_month.get())
            base_name += f"_{month_num}"
        
        if self.target_year.get() != "All":
            base_name += f"_{self.target_year.get()}"
        
        if self.target_month.get() == "All" and self.target_year.get() == "All":
            current_date = datetime.now().strftime("%Y_%m_%d")
            base_name += f"_all_data_{current_date}"
        
        return f"{base_name}.csv"
    
    def update_filename(self):
        if not self.custom_output_path.get():
            new_filename = self.generate_filename()
            self.output_filename.set(new_filename)
            self.log(f"Auto-updated filename to: {new_filename}")
    
    def choose_output_location(self):
        try:
            filename = filedialog.asksaveasfilename(
                title="Choose where to save the attendance report",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=self.output_filename.get()
            )
            if filename:
                self.custom_output_path.set(filename)
                self.output_filename.set(os.path.basename(filename))
                directory = os.path.dirname(filename)
                self.output_location.set(f"Custom location: {directory}")
                self.log(f"Output location changed to: {filename}")
            else:
                self.log("Output location selection cancelled")
        except Exception as e:
            self.log(f"Error choosing output location: {e}")
            messagebox.showerror("Error", f"Error selecting output location: {e}")
    
    def preview_files(self):
        if not self.folder_path.get():
            messagebox.showerror("Error", "Please select a folder first!")
            return
            
        pattern = os.path.join(self.folder_path.get(), self.file_pattern.get())
        all_files = glob.glob(pattern)
        
        tournament_files = []
        skipped_files = []
        
        for file_path in all_files:
            if self.is_tournament_file(file_path):
                tournament_files.append(file_path)
            else:
                skipped_files.append(file_path)
        
        self.log(f"Searching for files matching: {pattern}")
        self.log(f"Found {len(all_files)} files total, {len(tournament_files)} valid tournament files")
        
        if skipped_files:
            self.log(f"Skipped {len(skipped_files)} non-tournament files:")
            for path in skipped_files[:5]:
                self.log(f"  - {os.path.basename(path)}")
            if len(skipped_files) > 5:
                self.log(f"  ... and {len(skipped_files) - 5} more")
        
        if not tournament_files:
            self.log("No valid tournament files found!")
            return
            
        self.log(f"\nValid tournament files:")
        for path in tournament_files:
            self.log(f"  - {os.path.basename(path)}")
            
        self.log("\nTournament dates:")
        for path in tournament_files:
            try:
                tree = ET.parse(path)
                root = tree.getroot()
                start_date_elem = root.find('data/startdate')
                if start_date_elem is not None:
                    self.log(f"  {os.path.basename(path)}: {start_date_elem.text}")
            except Exception as e:
                self.log(f"  Error reading {os.path.basename(path)}: {e}")
                
    def generate_report(self):
        thread = threading.Thread(target=self.generate_report_thread)
        thread.daemon = True
        thread.start()
        
    def get_month_number(self, month_name):
        months = {
            'January': 1, 'February': 2, 'March': 3, 'April': 4,
            'May': 5, 'June': 6, 'July': 7, 'August': 8,
            'September': 9, 'October': 10, 'November': 11, 'December': 12
        }
        return months.get(month_name)
        
    def is_tournament_file(self, file_path):
        try:
            if not file_path.lower().endswith(('.xml', '.tdf')):
                return False
                
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            if root.tag != 'tournament':
                return False
                
            data_section = root.find('data')
            players_section = root.find('players')
            
            return data_section is not None and players_section is not None
            
        except ET.ParseError:
            return False
        except Exception:
            return False

    def parse_tournament_file(self, file_path):
        try:
            if not self.is_tournament_file(file_path):
                self.log(f"Skipping {os.path.basename(file_path)} - not a tournament file")
                return []
                
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            start_date_elem = root.find('data/startdate')
            if start_date_elem is None:
                self.log(f"Warning: No startdate found in {os.path.basename(file_path)}")
                return []
            
            start_date = start_date_elem.text
            tournament_date = datetime.strptime(start_date, '%m/%d/%Y')
            
            players = []
            players_section = root.find('players')
            if players_section is None:
                self.log(f"Warning: No players section found in {os.path.basename(file_path)}")
                return []
            
            for player in players_section.findall('player'):
                user_id = player.get('userid')
                
                first_name_elem = player.find('firstname')
                last_name_elem = player.find('lastname')
                birth_date_elem = player.find('birthdate')
                
                if first_name_elem is None or last_name_elem is None or birth_date_elem is None:
                    continue
                
                first_name = first_name_elem.text
                last_name = last_name_elem.text
                birth_date = birth_date_elem.text
                
                birth_year = datetime.strptime(birth_date, '%m/%d/%Y').year
                
                players.append({
                    'user_id': user_id,
                    'first_name': first_name,
                    'last_name': last_name,
                    'birth_year': birth_year,
                    'tournament_date': tournament_date
                })
            
            self.log(f"Parsed {len(players)} players from {os.path.basename(file_path)} ({start_date})")
            return players
        except Exception as e:
            self.log(f"Skipping {os.path.basename(file_path)} - parsing error: {e}")
            return []
            
    def analyze_attendance(self):
        if not self.folder_path.get():
            messagebox.showerror("Error", "Please select a folder first!")
            return None
            
        pattern = os.path.join(self.folder_path.get(), self.file_pattern.get())
        all_files = glob.glob(pattern)
        
        file_paths = [f for f in all_files if self.is_tournament_file(f)]
        
        if not file_paths:
            messagebox.showerror("Error", "No valid tournament files found!")
            return None
            
        self.log(f"Processing {len(file_paths)} valid tournament files...")
        
        target_month = None
        target_year = None
        
        if self.target_month.get() != "All":
            target_month = self.get_month_number(self.target_month.get())
            
        if self.target_year.get() != "All":
            target_year = int(self.target_year.get())
            
        player_attendance = defaultdict(lambda: {
            'first_name': '',
            'last_name': '',
            'birth_year': 0,
            'attendance_count': 0
        })
        
        for file_path in file_paths:
            players = self.parse_tournament_file(file_path)
            
            for player in players:
                tournament_date = player['tournament_date']
                
                if target_month and tournament_date.month != target_month:
                    continue
                if target_year and tournament_date.year != target_year:
                    continue
                
                user_id = player['user_id']
                player_attendance[user_id]['first_name'] = player['first_name']
                player_attendance[user_id]['last_name'] = player['last_name']
                player_attendance[user_id]['birth_year'] = player['birth_year']
                player_attendance[user_id]['attendance_count'] += 1
                
        return player_attendance
        
    def write_csv(self, player_attendance, output_path):
        try:
            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['player_id', 'first_name', 'last_name', 'birth_year', 'attendance_count']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                writer.writeheader()
                for player_id, data in sorted(player_attendance.items(), 
                                            key=lambda x: x[1]['attendance_count'], reverse=True):
                    writer.writerow({
                        'player_id': player_id,
                        'first_name': data['first_name'],
                        'last_name': data['last_name'],
                        'birth_year': data['birth_year'],
                        'attendance_count': data['attendance_count']
                    })
            return True, None
        except PermissionError as e:
            return False, f"Permission denied. Cannot write to this location.\n\nTry:\n1. Run as administrator, or\n2. Choose a different output location like Documents or Desktop"
        except Exception as e:
            return False, f"Error writing file: {e}"
    
    def get_safe_output_path(self, preferred_path):
        try:
            test_file = preferred_path.replace('.csv', '_test.tmp')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            return preferred_path
        except (PermissionError, OSError):
            try:
                import winreg
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                   r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
                documents_dir = winreg.QueryValueEx(key, "Personal")[0]
                winreg.CloseKey(key)
                filename = os.path.basename(preferred_path)
                return os.path.join(documents_dir, filename)
            except:
                filename = os.path.basename(preferred_path)
                return os.path.join(os.path.expanduser("~"), filename)
    
    def get_output_path(self):
        if self.custom_output_path.get():
            return self.custom_output_path.get()
        else:
            return os.path.join(self.folder_path.get(), self.output_filename.get())
            
    def generate_report_thread(self):
        try:
            self.progress.start()
            
            player_attendance = self.analyze_attendance()
            if player_attendance is None:
                return
                
            preferred_output_path = self.get_output_path()
            safe_output_path = self.get_safe_output_path(preferred_output_path)
            
            success, error_msg = self.write_csv(player_attendance, safe_output_path)
            
            if not success:
                self.log(f"Failed to write to {safe_output_path}")
                self.log(f"Error: {error_msg}")
                
                messagebox.showerror("Permission Error", 
                                   f"Cannot write to the selected location.\n\n"
                                   f"Error: {error_msg}\n\n"
                                   f"Please click 'Choose Location' to select a different folder "
                                   f"(like Documents or Desktop) and try again.")
                return
            
            if safe_output_path != preferred_output_path:
                self.log(f"Used alternative location: {safe_output_path}")
                
            if player_attendance:
                attendance_counts = [data['attendance_count'] for data in player_attendance.values()]
                avg_attendance = sum(attendance_counts) / len(attendance_counts)
                
                self.log(f"\n=== REPORT GENERATED ===")
                self.log(f"Output file: {safe_output_path}")
                self.log(f"Total unique players: {len(player_attendance)}")
                self.log(f"Average attendance per player: {avg_attendance:.1f}")
                self.log(f"Max attendance: {max(attendance_counts)}")
                self.log(f"Min attendance: {min(attendance_counts)}")
                
                sorted_players = sorted(player_attendance.items(), 
                                      key=lambda x: x[1]['attendance_count'], reverse=True)
                self.log(f"\nTop 5 attendees:")
                for i, (player_id, data) in enumerate(sorted_players[:5]):
                    self.log(f"  {i+1}. {data['first_name']} {data['last_name']} - {data['attendance_count']} tournaments")
                
                success_msg = (f"Report generated successfully!\n\n"
                             f"File: {os.path.basename(safe_output_path)}\n"
                             f"Location: {os.path.dirname(safe_output_path)}\n"
                             f"Players: {len(player_attendance)}\n"
                             f"Average attendance: {avg_attendance:.1f}")
                
                if safe_output_path != preferred_output_path:
                    success_msg += f"\n\nNote: File was saved to an alternative location due to permissions."
                
                messagebox.showinfo("Success", success_msg)
            else:
                self.log("No players found matching the criteria!")
                messagebox.showwarning("Warning", "No players found matching the selected criteria!")
                
        except Exception as e:
            self.log(f"Error generating report: {e}")
            messagebox.showerror("Error", f"Failed to generate report: {e}")
        finally:
            self.progress.stop()

def main():
    root = tk.Tk()
    app = TournamentAttendanceGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

def run_app():
    main()