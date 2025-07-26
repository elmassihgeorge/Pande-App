#!/usr/bin/env python3
"""
Tournament Attendance Analyzer Launcher
"""

import sys
import os

def main():
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)
        sys.path.insert(0, script_dir)
        
        import tournament_gui
        tournament_gui.run_app()
        
    except FileNotFoundError:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing File", 
            f"Cannot find tournament_gui.py in the same folder as this launcher.\n\n"
            f"Please ensure both files are in the same directory:\n"
            f"- launcher.py\n"
            f"- tournament_gui.py\n\n"
            f"Current directory: {os.getcwd()}"
        )
        
    except Exception as e:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Error", 
            f"An error occurred while starting the application:\n\n{e}\n\n"
            f"Try running tournament_gui.py directly."
        )

if __name__ == "__main__":
    main()