import os
import pandas as pd
from datetime import datetime

# Folder path containing multiple test case files
test_case_folder = r"E:\Priojecttree\Automation Testing\Git\ALM\Test Case Files"
output_folder = r"E:\Priojecttree\Automation Testing\Git\ALM\Excel Report"

# Initialize an empty DataFrame to store combined data
combined_data = pd.DataFrame()

# Iterate through all Excel files in the folder
for file_name in os.listdir(test_case_folder):
    if file_name.endswith(".xlsx"):  # Process only Excel files
        file_path = os.path.join(test_case_folder, file_name)

        try:
            # Read the test case file
            test_case_data = pd.read_excel(file_path)

            # Add a column to identify the source file
            test_case_data['Source File'] = file_name

            # Append the data to the combined DataFrame
            combined_data = pd.concat([combined_data, test_case_data], ignore_index=True)

        except Exception as e:
            print(f"Error reading file {file_name}: {e}")

# Save the combined data to a new Excel file
if not combined_data.empty:
    current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_file_path = os.path.join(output_folder, f"ALM_Test_Results_{current_datetime}.xlsx")

    combined_data.to_excel(output_file_path, index=False)
    print(f"Test results saved to: {output_file_path}")
else:
    print("No test case data found in the folder.")
