import pandas as pd


# File paths
testcase_file_path = r"E:\Priojecttree\Automation Testing\Git\ALM\utilities\Test_Case.xlsx"
data_file_path = r"E:\Priojecttree\Automation Testing\Git\ALM\utilities\ALM_Data.xlsx"

def get_matched_test_cases(testcase_file_path, data_file_path):
    try:
        # Load the Test Case file (File 1)
        testcase_df = pd.read_excel(testcase_file_path, sheet_name="Test Case")

        # Filter out rows where 'Execute' is 'Yes'
        testcase_filtered_df = testcase_df[testcase_df['Execute'] == 'Yes']

        # List of all unique modules from the Test Case file
        modules_to_process = testcase_filtered_df['Module'].unique()

        # Initialize a list to store the merged data
        merged_data = []

        # Process each module in the filtered Test Case file
        for module in modules_to_process:
            try:
                # Load the corresponding sheet from the Data file (File 2)
                module_data = pd.read_excel(data_file_path, sheet_name=module, header=0)

                # Get the MTC values that are relevant to this module
                mtc_values = testcase_filtered_df[testcase_filtered_df['Module'] == module]['MTC'].tolist()

                # Filter the data from this module sheet based on the MTC values
                filtered_module_data = module_data[module_data['MTC'].isin(mtc_values)]

                # Merge the filtered module data with the relevant testcase data (testcase_df first)
                merged_module_data = pd.merge(testcase_filtered_df[testcase_filtered_df['Module'] == module],filtered_module_data,on='MTC', how='left')

                # Append the merged data to the list
                merged_data.append(merged_module_data)

            except Exception as e:
                print(f"Error processing module {module}: {e}")

        # Combine all the filtered and merged data into one DataFrame
        final_merged_df = pd.concat(merged_data, ignore_index=True, sort=False)

        return final_merged_df

    except Exception as e:
        print(f"Error loading or processing Excel files: {e}")
        return None
# Call the merge function
final_merged_df = get_matched_test_cases(testcase_file_path, data_file_path)

# If the merge was successful, print the merged DataFrame
if final_merged_df is not None:
    print(final_merged_df)
else:
    print("Failed to process the data.")