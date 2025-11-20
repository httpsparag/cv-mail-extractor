#!/usr/bin/env python
"""
CV Email Extractor - Startup Script
Handles virtual environment setup and application launch
"""

import os
import sys
import subprocess
import webbrowser
import time
from pathlib import Path

def run_command(command):
    """Run a shell command"""
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        return result.returncode == 0, result.stdout, result.stderr
    except Exception as e:
        return False, "", str(e)

def main():
    print("\n" + "="*50)
    print("  CV Email Extractor - Web UI")
    print("="*50 + "\n")
    
    project_dir = Path(__file__).parent
    venv_dir = project_dir / ".venv"
    
    # Create virtual environment if it doesn't exist
    if not venv_dir.exists():
        print("Creating virtual environment...")
        success, stdout, stderr = run_command(f'"{sys.executable}" -m venv "{venv_dir}"')
        if not success:
            print(f"❌ Error creating virtual environment: {stderr}")
            return 1
        print("✓ Virtual environment created")
    
    # Determine the Python executable in the virtual environment
    if sys.platform == "win32":
        python_exe = venv_dir / "Scripts" / "python.exe"
        pip_exe = venv_dir / "Scripts" / "pip.exe"
    else:
        python_exe = venv_dir / "bin" / "python"
        pip_exe = venv_dir / "bin" / "pip"
    
    # Install requirements
    print("Installing dependencies...")
    requirements_file = project_dir / "requirements.txt"
    success, stdout, stderr = run_command(f'"{pip_exe}" install -r "{requirements_file}" -q')
    if not success:
        print(f"⚠ Warning: Some dependencies may not have installed properly")
        print(f"Error output: {stderr}")
    else:
        print("✓ Dependencies installed")
    
    # Start the application
    print("\n" + "="*50)
    print("  Starting Application")
    print("="*50 + "\n")
    
    print("Opening http://localhost:5000 in your browser...")
    print("Press CTRL+C to stop the server\n")
    
    # Try to open browser after a short delay
    try:
        time.sleep(2)
        webbrowser.open('http://localhost:5000')
    except Exception as e:
        print(f"Note: Could not automatically open browser: {e}")
        print("Please manually navigate to http://localhost:5000")
    
    # Run the Flask app
    app_file = project_dir / "app.py"
    os.chdir(project_dir)
    
    try:
        if sys.platform == "win32":
            os.system(f'"{python_exe}" "{app_file}"')
        else:
            subprocess.run([str(python_exe), str(app_file)])
    except KeyboardInterrupt:
        print("\n\nApplication stopped.")
    except Exception as e:
        print(f"\n❌ Error running application: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
