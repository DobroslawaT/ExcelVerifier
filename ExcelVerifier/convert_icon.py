"""
Convert PNG icon to ICO format for PyInstaller
"""
from PIL import Image
import os

# Input and output paths
png_path = r"C:\Users\dobro\Downloads\Ikona apka.png"
ico_path = os.path.join(os.path.dirname(__file__), "icon.ico")

try:
    # Open the PNG image
    img = Image.open(png_path)
    
    # Convert RGBA to RGB if necessary (ICO format may require RGB)
    if img.mode == 'RGBA':
        # Create a white background
        background = Image.new('RGB', img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[3])  # Use alpha channel as mask
        img = background
    elif img.mode != 'RGB':
        img = img.convert('RGB')
    
    # Resize to standard icon sizes if needed (ICO supports multiple sizes)
    # Standard size for icons is 256x256
    if img.size != (256, 256):
        img = img.resize((256, 256), Image.Resampling.LANCZOS)
    
    # Save as ICO
    img.save(ico_path, 'ICO')
    
    print(f"✓ Successfully converted PNG to ICO!")
    print(f"✓ Saved to: {ico_path}")
    print(f"\nYou can now use this command to build the exe:")
    print(f'  .\.venv\Scripts\pyinstaller --onefile --windowed --name "ExcelVerifier" --icon=icon.ico ExcelVerifier/main.py')
    
except FileNotFoundError:
    print(f"✗ Error: PNG file not found at {png_path}")
    print(f"  Please check the path and try again")
except Exception as e:
    print(f"✗ Error converting image: {e}")
    print(f"  Make sure Pillow is installed: .\.venv\Scripts\pip install pillow")
