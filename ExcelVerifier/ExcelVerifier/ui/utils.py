from openpyxl.styles.colors import COLOR_INDEX

def hex6_from_rgb(rgb):
    """
    Safely extracts a 6-digit hex color from openpyxl values.
    Handles 'RRGGBB' and 'FFRRGGBB' formats.
    """
    if not rgb or not isinstance(rgb, str):
        return None
        
    # Openpyxl often returns ARGB like 'FFRRGGBB'.
    # We only want the last 6 characters (RRGGBB).
    hex6 = rgb[-6:]
    
    # Check if it is a valid hex string
    try:
        int(hex6, 16)
    except Exception:
        return None

    # Treat pure default black (000000) as "no color" / default
    if hex6 == '000000':
        return None
        
    return '#' + hex6

def resolve_excel_color(cell_color):
    """
    Resolves a QColor-compatible hex string from an openpyxl Color object.
    Handles RGB, Indexed, and Theme colors (partially).
    """
    if cell_color is None:
        return None

    # 1. Try RGB first (Most common for custom formatting)
    # Openpyxl Color objects often store the value in the .rgb attribute
    rgb = getattr(cell_color, 'rgb', None)
    hexcol = hex6_from_rgb(rgb)
    if hexcol:
        return hexcol

    # 2. Try Indexed Colors (Legacy Excel colors, e.g., Index 64)
    idx = getattr(cell_color, 'indexed', None)
    if idx is not None:
        try:
            # COLOR_INDEX is a list provided by openpyxl mapping index -> ARGB
            if 0 <= int(idx) < len(COLOR_INDEX):
                mapped = COLOR_INDEX[int(idx)]
                # Recursively clean the result (it usually comes back as FFRRGGBB)
                return hex6_from_rgb(mapped)
        except Exception:
            pass

    # 3. Theme colors are complex and require the workbook theme XML.
    # For a basic app, we usually skip them or default to None.
    return None