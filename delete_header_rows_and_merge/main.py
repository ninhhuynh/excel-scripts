import os
import pandas as pd

DATA_FOLDER_PATH = "Data"
HEADER_ROW_INDEX = 5
DEFAULT_DATA_TYPE = str
HEADER_COLUMN_NAMES = ["Product Name", "Seller SKU", "SKU ID", "URL", "Teasing Visitors", "Reminders",
                       "FS Visitors", "Add to Cart Visitors", "Revenue", "Orders", "Buyers", "Conversion", "Unit Sold", "Price"]

all_files = os.listdir(DATA_FOLDER_PATH)

if not len(all_files) > 0:
    raise ValueError("No file found")

print(f"Found {len(all_files)} files")

dfs: list[pd.DataFrame] = []
for f in all_files:
    print(f"Reading: {f}")
    read_df = pd.read_excel(os.path.join(
        DATA_FOLDER_PATH, f), header=HEADER_ROW_INDEX, dtype=DEFAULT_DATA_TYPE)

    if not (isinstance(read_df, pd.DataFrame) and len(read_df.index) > 0):
        raise ValueError(f"Cannot read data or file is empty: {f}")

    for column_name in read_df.columns:
        if column_name not in HEADER_COLUMN_NAMES:
            raise ValueError(
                f"Column names not detected: {column_name} for {f}")

    print("df read:")
    print(read_df)

    dfs.append(read_df)

print('Concatenating files...')
concat_df = pd.concat(dfs)

with pd.ExcelWriter('output.xlsx') as writer:
    print("writing to new file...")
    concat_df.to_excel(writer, index=False)
    print("completed your data will be in output.xlsx! Press any key to exit...")
    input()
