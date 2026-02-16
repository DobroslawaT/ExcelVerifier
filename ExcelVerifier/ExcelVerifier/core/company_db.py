import json
import os

from core.database_handler import DatabaseHandler
from config import DATABASE_FILE


def normalize_nip(value):
    if value is None:
        return ""
    return "".join(ch for ch in str(value) if ch.isdigit())


def load_company_db(file_path):
    db = DatabaseHandler(DATABASE_FILE)
    companies = db.get_companies()

    # If DB already has companies from migration, just use them
    if companies:
        return [{"name": c.get("name", ""), "nip": c.get("nip", "")} for c in companies]
    
    # Only migrate from legacy JSON if DB is completely empty
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as file_handle:
                data = json.load(file_handle)
            if isinstance(data, list):
                cleaned = []
                for item in data:
                    if not isinstance(item, dict):
                        continue
                    name = str(item.get("name", "")).strip()
                    nip = normalize_nip(item.get("nip", ""))
                    if name and nip:
                        cleaned.append({"name": name, "nip": nip})
                if cleaned:
                    # Add companies without deleting existing ones
                    for company in cleaned:
                        db.add_company(company['name'], company['nip'])
                    companies = db.get_companies()
        except Exception:
            return []

    return [{"name": c.get("name", ""), "nip": c.get("nip", "")} for c in companies]


def save_company_db(file_path, companies):
    try:
        print(f"[SAVE] Starting save with {len(companies)} companies")
        cleaned = []
        for item in companies:
            name = str(item.get("name", "")).strip()
            nip = normalize_nip(item.get("nip", ""))
            print(f"[SAVE] Processing: name={name}, nip={nip}")
            if name and nip:
                cleaned.append({"name": name, "nip": nip})
        print(f"[SAVE] Cleaned: {cleaned}")
        db = DatabaseHandler(DATABASE_FILE)
        db.replace_companies(cleaned)
        print(f"[SAVE] Success - saved {len(cleaned)} companies")
        return True
    except Exception as e:
        print(f"[SAVE] Error: {e}")
        import traceback
        traceback.print_exc()
        return False


def merge_companies(existing, new_items):
    merged = {item["nip"]: dict(item) for item in existing if item.get("nip")}
    for item in new_items:
        nip = item.get("nip")
        name = item.get("name")
        if not nip or not name:
            continue
        merged[nip] = {"name": name, "nip": nip}
    return list(merged.values())
