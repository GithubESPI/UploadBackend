import pandas as pd

# Load the uploaded Excel files
file_path_1 = '../../uploads/integrated_data.xlsx'
file_path_2 = '../../uploads/B-M1 MIFIM ALT - ALT Semestre 1.xlsx'
df_1 = pd.read_excel(file_path_1, header=None)
df_2 = pd.read_excel(file_path_2, header=None)

# Extract and display a few rows from each dataframe to locate the correct titles row
print("First few rows of file 1:")
print(df_1.head(10))  # Adjust the number of rows as needed

print("\nFirst few rows of file 2:")
print(df_2.head(10))  # Adjust the number of rows as needed

# Extract the titles from the specified cells
titles_row_1 = df_1.iloc[0, 2:22].tolist()  # C1 to V1
titles_row_2 = df_2.iloc[0, 2:22].tolist()  # Adjusted to fourth row (index 3)

# Remove nan values from the titles
titles_row_1_cleaned = [title for title in titles_row_1 if pd.notna(title)]
titles_row_2_cleaned = [title for title in titles_row_2 if pd.notna(title)]

print("Titles Row 1 (Cleaned):", titles_row_1_cleaned)
print("Titles Row 2 (Cleaned):", titles_row_2_cleaned)

# Check if the cleaned titles match
if titles_row_1_cleaned == titles_row_2_cleaned:
    print("The titles match.")
else:
    print("The titles do not match.")