# Tournament Attendance Analyzer

A simple desktop application for analyzing Pokemon tournament attendance from tournament files (.tdf format).

## Quick Setup

### 1. Download Files
- Download all files from GitHub to a folder on your computer
- You need: `tournament_gui.py`, `launcher.py`, and `run_tournament_app.bat`

### 2. Install Python (if not already installed)
- Go to https://python.org/downloads
- Download Python 3.6 or newer
- During installation, check "Add Python to PATH"

### 3. Create Desktop Shortcut
- Right-click on `run_tournament_app.bat`
- Select "Create shortcut"
- Move the shortcut to your desktop
- Rename it to "Tournament Analyzer" (optional)

## How to Use

### 1. Start the App
- Double-click the desktop shortcut
- The Tournament Attendance Analyzer window will open

### 2. Select Tournament Files
- Click "Browse" to choose the folder containing your tournament files (.tdf files)
- The app will automatically detect valid tournament files

### 3. Choose Time Period
- **Month**: Select specific month or "All" for entire year
- **Year**: Select specific year or "All" for all years
- Filename updates automatically based on your selection

### 4. Generate Report
- Click "Preview Files" to see what tournaments will be included (optional)
- Click "Generate Report" to create the attendance CSV
- If you get permission errors, click "Choose Location" to save elsewhere

### 5. View Results
- The CSV file contains: Player ID, First Name, Last Name, Birth Year, Attendance Count
- Players are sorted by attendance (highest first)
- File is automatically named with date filters (e.g., "tournament_attendance_6_2025.csv")

## Output Examples
- June 2025: `tournament_attendance_6_2025.csv`
- All of 2024: `tournament_attendance_2024.csv` 
- All data: `tournament_attendance_all_data_2025_07_26.csv`

## Troubleshooting

**App won't start?**
- Make sure Python is installed with "Add to PATH" checked
- Try running `tournament_gui.py` directly by double-clicking it

**Permission errors when saving?**
- Click "Choose Location" and save to Documents or Desktop folder
- Or right-click the app shortcut and "Run as administrator"

**No files found?**
- Check that tournament files have .tdf extension
- Use "Preview Files" to see what the app detected
- Verify the correct folder is selected

## Support
For issues or questions, email elmassihgeorge at gmail dot com