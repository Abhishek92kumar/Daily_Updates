import os
import json
import glob
import pandas as pd

# 1. Define the folder path containing your JSON files
folder_path = r"D:\json"

# 2. Find all JSON files in the folder automatically
json_pattern = os.path.join(folder_path, "*.json")
file_list = glob.glob(json_pattern)

all_classifications = []
total_expected = 0
total_processed = 0

# 3. Loop through every file and extract the data
for file_path in file_list:
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
        # Aggregate verification counts
        if "verification" in data:
            total_expected += data["verification"].get("expected_question_count", 0)
            total_processed += data["verification"].get("processed_question_count", 0)
            
        # Collect all classification dictionaries
        if "classifications" in data:
            all_classifications.extend(data["classifications"])

# 4. Convert the aggregated data into a Pandas DataFrame
df = pd.DataFrame(all_classifications)

# 5. Expand the list of tags so each tag gets its own row
df = df.explode('tags')

# 6. Blank out repeated 'q_id' and 'primary_category' values to match your requested Excel layout
# This keeps the ID and Category only on the first row of each new question
df.loc[df.duplicated(subset=['q_id', 'primary_category']), ['q_id', 'primary_category']] = ""

# Add an empty 'classifications' column at the far left to match your header requirement
# df.insert(0, 'classifications', "")

# Reorder columns to ensure they match exactly: classifications, q_id, primary_category, tags
df = df[['q_id', 'primary_category', 'tags']]

# 7. Save to Excel
output_excel = os.path.join(folder_path, "merged_output.xlsx")

# index=False prevents pandas from writing row numbers into the file
df.to_excel(output_excel, index=False)

print(f"Successfully processed {len(file_list)} files.")
print(f"Total processed questions: {total_processed}")
print(f"Excel file created at: {output_excel}")
