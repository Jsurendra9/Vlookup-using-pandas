import pandas as pd

pm = pd.read_excel("pm.xlsx", sheet_name="main")
pms = pd.read_excel("pm.xlsx", sheet_name="state")

pm["DOD"] = pm["DOD"].fillna("--")

vlookup = pd.merge(pm, pms[['Name']], on='Name', how="inner")
vlookup.to_excel("Result.xlsx", sheet_name="Vlook up", index=False)

pms["Not in Office"] = pm["Age"] - pm["Years in Office"]

vni = pd.merge(pm, pms, on='Name', how="inner")

with pd.ExcelWriter("Result.xlsx", engine="openpyxl", mode="a") as writer:
   vni.to_excel(writer, sheet_name="Inner", index=False)

vno = pd.merge(pm, pms, on='Name', how="outer")

with pd.ExcelWriter("Result.xlsx",engine="openpyxl", mode="a") as writer:
   vno.to_excel(writer, sheet_name="Outer", index=False)

vnl = pd.merge(pm, pms, on='Name', how="left")

with pd.ExcelWriter("Result.xlsx", engine="openpyxl", mode="a") as writer:
   vnl.to_excel(writer, sheet_name="Left", index=False)

vnr = pd.merge(pm, pms, on='Name', how="right")

with pd.ExcelWriter("Result.xlsx", engine="openpyxl", mode="a") as writer:
   vnr.to_excel(writer, sheet_name="Right", index=False)

print("File is created")

print("""

- **'inner'**: Retains rows with keys present in both DataFrames, keeping only the common values based on the specified key(s).
- **'outer'**: Combines all rows from both DataFrames, merging where keys match and filling NaN for missing values in non-matching rows.
- **'left'**: Keeps all rows from the left DataFrame, merging matched rows from the right DataFrame and filling NaN for non-matches.
- **'right'**: Preserves all rows from the right DataFrame, merging matched rows from the left DataFrame and filling NaN for non-matches in the left DataFrame. 
- These options in the `how` parameter define how the data is merged, accommodating various scenarios and analytical needs based on the available data in the DataFrames. """)
