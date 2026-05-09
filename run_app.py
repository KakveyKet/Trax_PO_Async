import os
import sys
import streamlit.web.cli as stcli

# --- FORCED IMPORTS FOR PYINSTALLER ---
import pandas
import PyPDF2
import docx
import openpyxl
import sqlite3
# --------------------------------------

if __name__ == "__main__":
    try:
        # Determine if running in a bundle
        if getattr(sys, 'frozen', False):
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        # Get the absolute path to your app.py
        script_path = os.path.join(application_path, 'app.py')
        
        # CRASH CHECK: Make sure app.py was actually bundled!
        if not os.path.exists(script_path):
            print(f"CRITICAL ERROR: Cannot find app.py at {script_path}")
            print("Did you forget the --add-data flag during the build?")
            input("Press Enter to close...")
            sys.exit(1)
        
        # Simulate the "streamlit run app.py" command
        sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
        
        # Start the Streamlit server
        sys.exit(stcli.main())
        
    except Exception as e:
        # If ANYTHING goes wrong, print it and wait
        print(f"FATAL SYSTEM ERROR: {e}")
        input("Press Enter to close...")