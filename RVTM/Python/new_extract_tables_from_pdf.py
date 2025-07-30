import pdfplumber
import pandas as pd
import re
import os

# === CONFIG ===
INPUT_PDF  = "C:\\Users\\arora\\OneDrive - John Holland Group\\Github Repos\\Python Projects\\ARC-0431 Copy.pdf"
OUTPUT_XL = "C:\\Users\\arora\\OneDrive - John Holland Group\\Github Repos\\Python Projects\\ARC-0431 Copy2.xlsx"

# === UTIL: keep every row to the same length ===
def normalize_row(row, target_len):
    cells = list(row)
    # pad short rows
    if len(cells) < target_len:
        cells += [""] * (target_len - len(cells))
    # merge extra columns into last cell
    elif len(cells) > target_len:
        extras = cells[target_len-1:]
        merged = "\n".join(str(x).strip() for x in extras if x and x.strip())
        cells = cells[:target_len-1] + [merged]
    return cells

# === MAIN EXTRACTION ===
if not os.path.exists(INPUT_PDF):
    raise FileNotFoundError(f"Cannot find PDF: {INPUT_PDF}")

header     = None
target_len = None
all_rows   = []

with pdfplumber.open(INPUT_PDF) as pdf:
    for page in pdf.pages:
        table = page.extract_table({
            "vertical_strategy":   "lines",
            "horizontal_strategy": "lines"
        })
        if not table:
            continue

        # 1) On the first page that has it, detect your real header row
        if header is None:
            for i, r in enumerate(table):
                if r[0] and r[0].strip().startswith("RPV"):
                    header     = r
                    target_len = len(r)
                    break
        if header is None:
            continue

        # 2) Find where data starts on *this* page (just after the header)
        start_idx = 0
        for i, r in enumerate(table):
            if r[0] and r[0].strip().startswith("RPV"):
                start_idx = i + 1
                break

        # 3) Collect **all** rows until you hit a footer or blank row
        for r in table[start_idx:]:
            # skip repeated header row
            if r[0] and r[0].strip().startswith("RPV"):
                continue
            # skip footer rows like "Page 12 of 50"
            if any(cell and re.search(r"Page \d+ of \d+", cell) for cell in r):
                continue
            # skip completely empty rows
            if all((not cell or not cell.strip()) for cell in r):
                continue
            # normalize length and save
            all_rows.append(normalize_row(r, target_len))

# === BUILD & CLEAN DATAFRAME ===
# 4) Turn into DataFrame in one go
clean_cols = [c.strip().replace("\n", " ") for c in header]
df = pd.DataFrame(all_rows, columns=clean_cols)

# 5) Drop the stray “Unnamed” column (the vertical separator)
df = df.loc[:, [c for c in df.columns if c and not c.startswith("Unnamed")]]

# 6) Reorder so “Design Package” & “Compliance Status” are adjacent
cols = df.columns.tolist()
if "Design Package" in cols and "Compliance Status" in cols:
    cols.remove("Design Package")
    cols.remove("Compliance Status")
    df = df[cols[:5] + ["Design Package", "Compliance Status"] + cols[5:]]

# === EXPORT TO EXCEL ===
df.to_excel(OUTPUT_XL, index=False)
print(f"✔ Clean table saved to {OUTPUT_XL}")
