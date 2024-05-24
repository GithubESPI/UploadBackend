import pandas as pd

# Load the uploaded Excel file
file_path = '../../uploads/integrated_data.xlsx'
df = pd.read_excel(file_path, header=None)

# Extract the titles from the specified cells
titles_row = df.iloc[0, 2:22].tolist()  # C1 to V1

# Extract the data for each student from the specified cells
data_rows = df.iloc[1:, 2:22].values.tolist()  # C2 to V2 for each student

# Print the titles and the data for each student
print("Titles:", titles_row)
for i, data_row in enumerate(data_rows, start=1):
    print(f"Data row {i}:", data_row)
