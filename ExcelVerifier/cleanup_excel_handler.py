# Read the file
with open('ExcelVerifier\\core\\excel_handler.py', 'r') as f:
    lines = f.readlines()

# Keep lines 1-860 and lines 1138 onwards
# Lines are 1-indexed in output but 0-indexed in arrays
cleaned_lines = lines[:860] + lines[1137:]

# Write back
with open('ExcelVerifier\\core\\excel_handler.py', 'w') as f:
    f.writelines(cleaned_lines)

print("✓ File cleaned! Removed lines 861-1137")
print(f"  Original: {len(lines)} lines")
print(f"  Cleaned: {len(cleaned_lines)} lines")
