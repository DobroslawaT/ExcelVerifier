"""
Build ExcelVerifier into standalone .exe with installer
Run: .\.venv\Scripts\python build.py
"""
import subprocess
import os
import sys
from pathlib import Path

print("=" * 60)
print("ExcelVerifier Build Script")
print("=" * 60)

# Step 1: Convert PNG to ICO
print("\n[Step 1/3] Converting icon PNG to ICO format...")
try:
    subprocess.run([sys.executable, "convert_icon.py"], check=True)
except subprocess.CalledProcessError:
    print("✗ Icon conversion failed. Continuing without custom icon...")

# Step 2: Install/Verify PyInstaller
print("\n[Step 2/3] Checking PyInstaller installation...")
try:
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "pyinstaller"],
        capture_output=True,
        check=True
    )
    print("✓ PyInstaller is ready")
except subprocess.CalledProcessError:
    print("✗ Failed to install PyInstaller")
    sys.exit(1)

# Step 3: Build the executable
print("\n[Step 3/3] Building ExcelVerifier.exe...")
print("This may take a few minutes...")

icon_arg = "--icon=icon.ico" if os.path.exists("icon.ico") else ""

pyinstaller_cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--windowed",
    "--name=ExcelVerifier",
    "--add-data=ExcelVerifier/ui;ui",
    "--add-data=ExcelVerifier/core;core",
    "--add-data=ExcelVerifier/config.py;.",
    "--collect-all=openpyxl",
    "--collect-all=google.generativeai",
]

if icon_arg:
    pyinstaller_cmd.append(icon_arg)

pyinstaller_cmd.append("ExcelVerifier/main.py")

try:
    subprocess.run(pyinstaller_cmd, check=True)
    
    exe_path = Path("dist/ExcelVerifier.exe")
    if exe_path.exists():
        print("\n" + "=" * 60)
        print("✓ BUILD SUCCESSFUL!")
        print("=" * 60)
        print(f"\n📦 Your executable is ready:")
        print(f"   Location: {exe_path.absolute()}")
        print(f"   Size: {exe_path.stat().st_size / (1024*1024):.1f} MB")
        print(f"\n✓ You can now distribute ExcelVerifier.exe to users!")
        print(f"  Users simply need to run the .exe file - no Python required!")
    else:
        print("✗ Build completed but exe not found")
        sys.exit(1)
        
except subprocess.CalledProcessError as e:
    print(f"\n✗ Build failed: {e}")
    sys.exit(1)
