import pdfplumber
import pandas as pd
import re


PDF_PATH    = r"C:\Users\arora\OneDrive - John Holland Group\Working_files\SMWLW\TSMO_PHYSICAL.pdf"
OUTPUT_XLSX = r"TSMO_physical.xlsx"


def tidy_text(cell: str) -> str:

    if not isinstance(cell, str):
        return cell
    no_cid  = re.sub(r'\(cid:\d+\)', '', cell)
    compact = re.sub(r'\s+', ' ', no_cid).strip()
    return re.sub(r'\.\s+(?=[A-Z])', '.\n', compact)

def make_unique(cols):
   
    counts = {}
    out = []
    for c in cols:
        base = c.strip()
        i = counts.get(base, 0)
        if i == 0:
            out.append(base)
        else:
            out.append(f"{base}_{i}")
        counts[base] = i+1
    return out

def extract_tables(pdf_path):
  
    chunks = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table_settings in (
                {"vertical_strategy":"lines","horizontal_strategy":"lines"},
                {"vertical_strategy":"lines","horizontal_strategy":"text"},
            ):
                raw_tables = page.extract_tables(table_settings=table_settings)
                if not raw_tables:
                    continue
                for tbl in raw_tables:
                    header, *rows = tbl
                    # skip completely empty tables
                    if not header or not any(header):
                        continue
                    # drop any row that exactly matches the header
                    data_rows = [r for r in rows if any(r)]  # remove fully empty
                    data_rows = [r for r in data_rows if not all(
                        (str(cell).strip()==str(hdr).strip()) 
                        for cell, hdr in zip(r, header)
                    )]
                    if not data_rows:
                        continue
                    df = pd.DataFrame(data_rows, columns=header)
                    df.columns = make_unique(df.columns.astype(str))
                    chunks.append(df)
                break
    if not chunks:
        raise RuntimeError("No tables found – check PDF_PATH")
    return chunks

def concat_and_clean(chunks):

    big = pd.concat(chunks, ignore_index=True)
    big.dropna(how="all", inplace=True)

 
    for col in big.columns:
        big[col] = big[col].apply(tidy_text)


    stripped = big.columns.str.replace(r'_\d+$','',regex=True).str.strip()
    big.columns = stripped
    out = pd.DataFrame()
    for col in pd.unique(stripped):
        group = [c for c in big.columns if c == col]
        if len(group) == 1:
            out[col] = big[group[0]]
        else:
            out[col] = big[group].bfill(axis=1).iloc[:,0]
    return out

def build_exact(df):

    desig_col = next(
        (c for c in df.columns if "Interface" in c and "Designation" in c),
        None
    )
    if desig_col is None:
        raise KeyError("Couldn't find the Interface Designation column")


    m = df[desig_col].astype(str).str.contains("SMW-PI-", na=False)
    sub = df[m].copy()
    sub[desig_col] = sub[desig_col].ffill()

    return sub.rename(columns={desig_col:"Interface Designation"})

def main():
    print("🔍 extracting table chunks...")
    chunks = extract_tables(PDF_PATH)
    print(f"   • got {len(chunks)} chunks")

    print("🧹 concatenating & cleaning...")
    cleaned = concat_and_clean(chunks)

    print("📋 building final sheet...")
    final = build_exact(cleaned)

    final.to_excel(OUTPUT_XLSX, index=False)
    print(f"✅ wrote {len(final)} rows to {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
