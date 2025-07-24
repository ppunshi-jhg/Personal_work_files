import pandas as pd
from functools import lru_cache

df = pd.read_excel("new_input.xlsx", sheet_name="Sheet1", dtype=str)
df["section"]= df["section"].fillna("").astype(str)
df["Primary Text"] = df["Primary Text"].fillna("").astype(str)

code_to_text = dict(zip(df["section"], df["Primary Text"]))

@lru_cache(maxsize=None)
def get_hierarchy_text(code: str) -> str:
    code = str(code)
    if not code or code not in code_to_text:
        return ""     
    parts: list = []

    for idx, ch in enumerate(code):
        if ch in (".", "-"):
            parent = code[:idx]
            if parent in code_to_text:
                parts.append(code_to_text[parent])
    parts.append(code_to_text[code])
    
    return "\n".join(parts)


df["Result"] = df["section"].apply(get_hierarchy_text)


df.to_excel("output_with_results_2.xlsx", index=False)


# Excel formula to do the same thing is:
# "=LET(
#   code,       [@section],
#   tblSec,     [section],
#   tblTxt,     [Primary Text],
#   seq,        SEQUENCE(LEN(code)),
#   sepPos,     FILTER(seq, (MID(code,seq,1)=""."") + (MID(code,seq,1)=""-"")),
#   parentCodes, LEFT(code, sepPos-1),
#   parentTexts, IFERROR( XLOOKUP(parentCodes, tblSec, tblTxt, """"), """" ),
#   thisText,    XLOOKUP(code,     tblSec, tblTxt, """"),
#   allTexts,    VSTACK(parentTexts, thisText),
#   TEXTJOIN(CHAR(10), TRUE, allTexts)
# )"

