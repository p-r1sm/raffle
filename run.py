#!/usr/bin/env python3
"""
Laabharti Card Generator - Launcher Script

This script provides a simple entry point for users to launch the card generator
application. It will attempt to install any missing dependencies and then
launch the GUI.
"""

import os
import sys
import subprocess
import platform

def check_dependencies():
    """Check if required packages are installed and install if missing."""
    try:
        import tkinter
        import pandas
        import docx
        import PIL
        return True
    except ImportError as e:
        missing_package = str(e).split("'")[1]
        print(f"Missing dependency: {missing_package}")
        
        try:
            if platform.system() == "Windows":
                python_executable = "python"
            else:
                python_executable = "python3"
                
            print(f"Attempting to install missing dependencies...")
            subprocess.check_call([python_executable, "-m", "pip", "install", "-r", "requirements.txt"])
            return True
        except Exception as install_error:
            print(f"Error installing dependencies: {install_error}")
            print("\nPlease manually install the required packages using:")
            print("pip install -r requirements.txt")
            
            if platform.system() != "Windows":
                print("\nOr try:")
                print("python3 -m pip install -r requirements.txt")
                
            input("\nPress Enter to exit...")
            return False

def main():
    """Main entry point of the application."""
    # Get the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Check and install dependencies if needed
    if not check_dependencies():
        return
    
    try:
        # Import and run the GUI
        from gui import CardGeneratorApp
        import tkinter as tk
        
        root = tk.Tk()
        app = CardGeneratorApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Error starting application: {e}")
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main() 