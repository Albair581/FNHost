from openpyxl import load_workbook

# Load workbook & sheet
wb = load_workbook("FN Research.xlsx")
sheet = wb.active

lob = []
fine = [{"org": sheet.title, "branches": []}]
# Loop through rows
for row in sheet.iter_rows(min_row=2, values_only=True, max_row=305, max_col=12):
    data = {
        "name": row[0] + f"（{row[1]}）",
        "address": row[2],
        "hours": row[7],
        "phone": row[4],
        "fax": row[5],
        "coordinates": row[3],
        "transportation": row[7],
        "parking": row[8],
        "accepted_items": row[9],
        "notes": row[10],
        "maps_url": row[11],
    }
    lob.append(data)

fine[0]["branches"] = lob

with open("fn-research.json", "w", encoding="utf-8") as f:
    import json

    json.dump(fine, f, ensure_ascii=False, indent=2)
