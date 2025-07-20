import pandas as pd

# Define the path to the Excel file
excel_file_path = 'Employee Sample Data - A.xlsx'
excel_output_path = 'cleaned_employee_data.xlsx'

print("--- Step 1: Data Exploration and Conversion to Pandas DataFrame ---")
try:
    # Load the Excel data into a Pandas DataFrame
    df = pd.read_excel(excel_file_path) # Changed to pd.read_excel
    print("DataFrame loaded successfully.")

print("\nInitial 5 rows of the DataFrame:")
print(df.head())

print("\nDataFrame Info (Data Types and Non-Null Counts):")
df.info()

print("\nDescriptive Statistics for Numerical Columns:")
print(df.describe())

print("\nMissing Values Count per Column:")
print(df.isnull().sum())

print("\n--- Step 2: Data Cleaning ---")
# Convert 'Hire Date' and 'Exit Date' to datetime objects
# Errors='coerce' will turn unparseable dates into NaT (Not a Time)
df['Hire Date'] = pd.to_datetime(df['Hire Date'], errors='coerce')
df['Exit Date'] = pd.to_datetime(df['Exit Date'], errors='coerce')

# Ensure 'Annual Salary' and 'Bonus %' are numeric
# In case they were loaded as objects due to non-numeric characters, convert them
df['Annual Salary'] = pd.to_numeric(df['Annual Salary'], errors='coerce')
df['Bonus %'] = pd.to_numeric(df['Bonus %'], errors='coerce')

# Fill any new NaNs created by 'coerce' for numeric columns if necessary (e.g., with 0 or mean)
# For this exercise, we'll assume the provided data is clean enough for these columns,
# but in a real-world scenario, you might fill NaNs or drop rows.

print("\nDataFrame Info after Date Conversion:")
df.info()

print("\n--- Step 3: Change the first 5 rows ---")
# Create new arbitrary values for the first 5 rows
new_data_for_first_5_rows = {
    'EEID': ['N0001', 'N0002', 'N0003', 'N0004', 'N0005'],
    'Full Name': ['New Employee One', 'New Employee Two', 'New Employee Three', 'New Employee Four', 'New Employee Five'],
    'Job Title': ['New Role 1', 'New Role 2', 'New Role 3', 'New Role 4', 'New Role 5'],
    'Department': ['New Dept A', 'New Dept B', 'New Dept A', 'New Dept C', 'New Dept B'],
    'Business Unit': ['New BU X', 'New BU Y', 'New BU X', 'New BU Z', 'New BU Y'],
    'Gender': ['Female', 'Male', 'Female', 'Male', 'Female'],
    'Ethnicity': ['Asian', 'Caucasian', 'Black', 'Latino', 'Asian'],
    'Age': [25, 30, 35, 40, 45],
    'Hire Date': pd.to_datetime(['2023-01-01', '2023-02-15', '2023-03-20', '2023-04-10', '2023-05-05']),
    'Annual Salary': [70000, 85000, 92000, 110000, 78000],
    'Bonus %': [0.05, 0.10, 0.08, 0.12, 0.06],
    'Country': ['Canada', 'Mexico', 'Germany', 'France', 'Spain'],
    'City': ['Toronto', 'Mexico City', 'Berlin', 'Paris', 'Madrid'],
    'Exit Date': [pd.NaT, pd.NaT, pd.NaT, pd.NaT, pd.NaT] # Use pd.NaT for missing dates
}

# Ensure the new data has the same columns as the original DataFrame
# Create a temporary DataFrame from the new data
temp_df = pd.DataFrame(new_data_for_first_5_rows)

# Update the first 5 rows of the original DataFrame
# Make sure to align columns correctly
for col in temp_df.columns:
    if col in df.columns:
        df.loc[0:4, col] = temp_df[col]

print("\nFirst 5 rows after modification:")
print(df.head())

print("\n--- Step 4: Print the row with the largest salary ---")
# Find the row with the largest Annual Salary
# .idxmax() returns the index of the first occurrence of the maximum value
row_largest_salary = df.loc[df['Annual Salary'].idxmax()]
print("\nRow with the largest Annual Salary:")
print(row_largest_salary)

print("\n--- Step 5: Group by department, and get the average age as well as average salary ---")
# Group by Department and calculate the mean of Age and Annual Salary
department_avg = df.groupby('Department').agg(
    Average_Age=('Age', 'mean'),
    Average_Salary=('Annual Salary', 'mean')
).round(2) # Round to 2 decimal places for better readability

print("\nAverage Age and Average Salary by Department:")
print(department_avg)

print("\n--- Step 6: Group by department+ethnicity, and find the maximum age, minimum age, and median salary ---")
# Group by Department and Ethnicity and calculate max age, min age, and median salary
department_ethnicity_stats = df.groupby(['Department', 'Ethnicity']).agg(
    Max_Age=('Age', 'max'),
    Min_Age=('Age', 'min'),
    Median_Salary=('Annual Salary', 'median')
).round(2) # Round to 2 decimal places

print("\nMax Age, Min Age, and Median Salary by Department and Ethnicity:")
print(department_ethnicity_stats)

print("\n--- Step 7: Save your work in a new Excel file ---")
# Save the modified DataFrame to a new Excel file
try:
    df.to_excel(excel_output_path, index=False)
    print(f"\nModified DataFrame successfully saved to '{excel_output_path}'")
except Exception as e:
    print(f"Error saving Excel file: {e}")

print("\n--- Task Completed ---")

