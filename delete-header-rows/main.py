import glob
import os
import pandas as pd

all_files = glob.glob(os.path.join("Data", "*.xls"))

if not len(all_files) > 0:
    raise ValueError("No file found")

dfs: list[pd.DataFrame] = []
for f in all_files:
    print(f"Reading: {f}")
    read_df = pd.read_excel(f, header=5, dtype=str)
    dfs.append(read_df)

for df in dfs:
    print("Checking...", df)
    for column_name in df.columns:
        if column_name not in ["Product Name", "Seller SKU", "SKU ID", "URL", "Teasing Visitors", "Reminders", "FS Visitors", "Add to Cart Visitors", "Revenue", "Orders", "Buyers", "Conversion", "Unit Sold", "Price"]:
            raise ValueError(f"Column names not detected: {column_name}")

print('Concatenating files...')
concat_df = pd.concat(dfs)

with pd.ExcelWriter('output.xlsx') as writer:
    print("writing to new file...")
    concat_df.to_excel(writer, index=False)
    print("completed your data will be in output.xlsx! Press any key to exit...")
    input()
