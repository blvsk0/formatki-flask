import os, pandas as pd, sys, json
fn = os.getenv("BASE_XLSX","baza.xlsx")
df = pd.read_excel(fn, sheet_name="Arkusz1", header=0, dtype=str).fillna('')
# Dostosuj te trzy wartości dokładnie tak jak testujesz w UI:
pion = "Technika"
gt = "1100 Gwoździe"
kw = "Zestawy gwoździ"

print("Looking for rows matching:")
print("pion repr:", repr(pion))
print("gt repr:", repr(gt))
print("kw repr:", repr(kw))
print("----\nSample of first 20 rows (cols GT, KW, PION, EAN, Nr. Art dostawcy, Punktor 1..5):\n")

cols_to_show = []
for c in ["GT","KW","PION","EAN","Nr. Art dostawcy"]:
    if c in df.columns: cols_to_show.append(c)
# add first 6 punktor cols (if exist)
punktory = [c for c in df.columns if str(c).strip().lower().startswith("punktor")]
cols_to_show += punktory[:6]

# print first 15 rows with repr
for i,row in df.head(15).iterrows():
    vals = {c: repr(str(row[c])) if c in row.index else "MISSING" for c in cols_to_show}
    print(i, json.dumps(vals, ensure_ascii=False))

print("\nNow filtering with exact match (lower+strip) and showing repr of selected rows:")
sel = df[
    (df["PION"].astype(str).str.strip().str.lower() == pion.strip().lower()) &
    (df["GT"].astype(str).str.strip().str.lower() == gt.strip().lower()) &
    (df["KW"].astype(str).str.strip().str.lower() == kw.strip().lower())
]
print("Matches count:", len(sel))
for i,row in sel.head(10).iterrows():
    vals = {c: repr(str(row[c])) if c in row.index else "MISSING" for c in cols_to_show + ["Podział"]}
    print(i, json.dumps(vals, ensure_ascii=False))
