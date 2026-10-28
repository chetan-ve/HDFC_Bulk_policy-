import glob
import os
import pandas as pd

# Target folder path
folder_path = r"assets\output\not_run_cases"
# folder_path = r"assets\output"

# Path ke andar ki saari Excel files ki list banana
file_pattern = os.path.join(folder_path, "remaining*.xlsx")
# file_pattern = os.path.join(folder_path, "output*.xlsx")

file_list = glob.glob(file_pattern)

# Saari files ko read karke merge karna
df_list = [pd.read_excel(file) for file in file_list]

if df_list:
    combined_df = pd.concat(df_list, ignore_index=True)

    # Output file ko ushi folder me save karna
    output_path = os.path.join(folder_path, "combined_remaining.xlsx")
    combined_df.to_excel(output_path, index=False)

    print(
        f"Successfully {len(file_list)} files combine ho gayi hain!\nSaved to: {output_path}"
    )
else:
    print("Koie matching file nahi mili. Path check karein.")