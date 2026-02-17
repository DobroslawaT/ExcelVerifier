"""
Build ExcelVerifier into standalone .exe files
Run: .\.venv\Scripts\python build.py
"""
import subprocess
import os
import sys
from pathlib import Path
import shutil

print("=" * 60)
print("ExcelVerifier Build Script")
print("=" * 60)

# Step 1: Convert PNG to ICO
print("\n[Step 1/4] Converting icon PNG to ICO format...")
try:
    subprocess.run([sys.executable, "convert_icon.py"], check=True)
except subprocess.CalledProcessError:
    print("✗ Icon conversion failed. Continuing without custom icon...")

# Step 2: Install/Verify PyInstaller
print("\n[Step 2/4] Checking PyInstaller installation...")
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

# Step 3: Build the main executable
print("\n[Step 3/4] Building ExcelVerifier.exe...")
print("This may take a few minutes...")

icon_arg = "--icon=icon.ico" if os.path.exists("icon.ico") else ""

pyinstaller_cmd = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm",
    "--onedir",
    "--windowed",
    "--name=ExcelVerifier",
    "--noupx",
    "--add-data=ExcelVerifier/ui;ui",
    "--add-data=ExcelVerifier/core;core",
    "--add-data=ExcelVerifier/config.py;.",
    # Only collect what we actually use
    "--hidden-import=PyQt5.QtCore",
    "--hidden-import=PyQt5.QtGui",
    "--hidden-import=PyQt5.QtWidgets",
    "--hidden-import=openpyxl",
    "--hidden-import=pandas",
    "--collect-all=google.generativeai",  # Collect ALL google.generativeai files
    "--collect-all=google.ai.generativelanguage",
    "--hidden-import=PIL",
    "--hidden-import=PIL.Image",
    "--hidden-import=win32crypt",
    "--hidden-import=win32api",
    # Email modules for googleapiclient
    "--hidden-import=email.mime",
    "--hidden-import=email.mime.multipart",
    "--hidden-import=email.mime.nonmultipart",
    "--hidden-import=email.mime.text",
    "--hidden-import=email.mime.base",
    # Exclude unnecessary heavy dependencies
    "--exclude-module=torch",
    "--exclude-module=tensorflow",
    "--exclude-module=matplotlib",
    "--exclude-module=scipy",
    "--exclude-module=notebook",
    "--exclude-module=jupyter",
    "--exclude-module=IPython",
]

if icon_arg:
    pyinstaller_cmd.append(icon_arg)

pyinstaller_cmd.append("ExcelVerifier/main.py")

try:
    subprocess.run(pyinstaller_cmd, check=True)
    print("✓ Main application built")
    
    # Manually copy PyQt5 plugins to avoid path encoding issues
    print("Copying PyQt5 plugins...")
    import site
    # Use parent directory's virtual environment site-packages
    script_dir = Path(__file__).parent
    venv_site_packages = script_dir.parent / ".venv" / "Lib" / "site-packages"
    site_packages = [str(venv_site_packages)] + site.getsitepackages()
    print(f"Using site-packages from: {venv_site_packages}")
    for site_dir in site_packages:
        qt5_plugins = Path(site_dir) / "PyQt5" / "Qt5" / "plugins"
        if qt5_plugins.exists():
            dist_plugins = Path("dist/ExcelVerifier/PyQt5/Qt5/plugins")
            if not dist_plugins.exists():
                shutil.copytree(qt5_plugins, dist_plugins)
                print(f"✓ Copied PyQt5 plugins")
            break
    
    # Manually copy ALL google modules to ensure complete package
    print("Copying all google modules...")
    google_modules = ['generativeai', 'ai', 'api', 'api_core', 'auth', 'protobuf', 'rpc', 'type', '_upb', 'oauth2', 'cloud', 'gapic', 'logging', 'longrunning']
    for site_dir in site_packages:
        google_dir = Path(site_dir) / "google"
        if google_dir.exists():
            dest_google = Path("dist/ExcelVerifier/_internal/google")
            dest_google.mkdir(parents=True, exist_ok=True)
            
            # Copy __init__.py
            init_file = google_dir / "__init__.py"
            if init_file.exists():
                shutil.copy(init_file, dest_google / "__init__.py")
            
            # Copy all necessary submodules
            for module in google_modules:
                module_dir = google_dir / module
                if module_dir.exists():
                    dest_module = dest_google / module
                    if dest_module.exists():
                        shutil.rmtree(dest_module)
                    shutil.copytree(module_dir, dest_module)
                    print(f"  ✓ Copied google.{module}")
            break
    
    # Copy all gRPC and related packages
    print("Copying gRPC and related packages...")
    extra_packages = ['grpc', 'grpc_status', 'proto', 'requests', 'urllib3', 'charset_normalizer', 'idna', 'certifi', 'googleapiclient', 'pydantic', 'pydantic_core', 'annotated_types', 'typing_inspection', 'tqdm', 'httplib2', 'uritemplate', 'pyparsing', 'pyasn1', 'pyasn1_modules', 'h11', 'sniffio', 'anyio', 'httpcore', 'httpx']
    for site_dir in site_packages:
        site_path = Path(site_dir)
        for pkg in extra_packages:
            pkg_dir = site_path / pkg
            if pkg_dir.exists():
                dest_pkg = Path("dist/ExcelVerifier/_internal") / pkg
                if dest_pkg.exists():
                    shutil.rmtree(dest_pkg)
                shutil.copytree(pkg_dir, dest_pkg)
                print(f"  ✓ Copied {pkg}")
        
        # Copy typing_extensions.py as a single file
        typing_ext = site_path / "typing_extensions.py"
        if typing_ext.exists():
            dest_typing = Path("dist/ExcelVerifier/_internal/typing_extensions.py")
            shutil.copy(typing_ext, dest_typing)
            print(f"  ✓ Copied typing_extensions.py")
        
        # Copy google_auth_httplib2.py as a single file
        google_auth_http = site_path / "google_auth_httplib2.py"
        if google_auth_http.exists():
            dest_auth = Path("dist/ExcelVerifier/_internal/google_auth_httplib2.py")
            shutil.copy(google_auth_http, dest_auth)
            print(f"  ✓ Copied google_auth_httplib2.py")
        
        # Copy six.py as a single file
        six_module = site_path / "six.py"
        if six_module.exists():
            dest_six = Path("dist/ExcelVerifier/_internal/six.py")
            shutil.copy(six_module, dest_six)
            print(f"  ✓ Copied six.py")
        break
    
except subprocess.CalledProcessError as e:
    print(f"✗ Main build failed: {e}")
    sys.exit(1)

# Step 4: Build the API configuration utility
print("\n[Step 4/4] Building configure_api.exe...")

configure_cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--console",
    "--name=configure_api",
    "--add-data=ExcelVerifier/config.py;.",
]

configure_cmd.append("configure_api.py")

try:
    subprocess.run(configure_cmd, check=True)
    
    # Copy configure_api.exe to the main dist folder
    config_exe = Path("dist/configure_api.exe")
    target_dir = Path("dist/ExcelVerifier")
    
    if config_exe.exists() and target_dir.exists():
        shutil.copy(config_exe, target_dir / "configure_api.exe")
        print("✓ API configuration utility built")
    
    exe_path = Path("dist/ExcelVerifier/ExcelVerifier.exe")
    if exe_path.exists():
        print("\n" + "=" * 60)
        print("✓ BUILD SUCCESSFUL!")
        print("=" * 60)
        print(f"\n📦 Your executables are ready:")
        print(f"   Main: {exe_path.absolute()}")
        print(f"   Config: {target_dir / 'configure_api.exe'}")
        print(f"\n✓ Ready for installer creation with Inno Setup!")
        print(f"  Run: iscc installer.iss")
    else:
        print("✗ Build completed but exe not found")
        sys.exit(1)
        
except subprocess.CalledProcessError as e:
    print(f"\n✗ Build failed: {e}")
    sys.exit(1)
