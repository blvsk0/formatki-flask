import os, pandas as pd, json
fn = os.getenv("BASE_XLSX","baza.xlsx")
print("Plik:", fn)
try:
    df = pd.read_excel(fn, sheet_name="Arkusz1", header=0, dtype=str)
except Exception as e:
    print("Błąd odczytu pliku:", e)
    raise SystemExit(1)
print()
print("Kolumny (order & names):")
for i,c in enumerate(df.columns.tolist()):
    print(f"{i}: '{c}'")
print()
print("Pierwsze 8 wierszy (kolumny 0..5):")
print(df.iloc[:8, :6].fillna('').to_string(index=True))
