"""
Configure API Key utility for ExcelVerifier installer
This script is called by the installer to set up the API key
"""
import sys
import os
import argparse

# Add parent directory to path to import config module
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from ExcelVerifier.config import set_gemini_api_key
except ImportError:
    # When running as part of built exe, the import path is different
    try:
        from config import set_gemini_api_key
    except ImportError:
        print("Error: Could not import config module")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description='Configure ExcelVerifier API Key')
    parser.add_argument('--api-key', type=str, help='Google Gemini API Key')
    
    args = parser.parse_args()
    
    # Skip if no API key provided
    if not args.api_key or args.api_key.strip() == "":
        print("No API key provided, skipping configuration")
        return 0
    
    try:
        # Save the API key using DPAPI encryption
        set_gemini_api_key(args.api_key.strip())
        print("API key configured successfully")
        return 0
    except Exception as e:
        print(f"Error configuring API key: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
