import pandas as pd

# File paths
test_case_file = r"E:\Priojecttree\Automation Testing\Git\ALM\utilities\Test_Case.xlsx"
data_file = r"E:\Priojecttree\Automation Testing\Git\ALM\utilities\ALM_Data.xlsx"

def get_matched_test_cases(test_case_file, data_file):

    # Load test cases where 'Execute' is 'Yes'
    test_case_df = pd.read_excel(test_case_file)
    filtered_test_case_df = test_case_df[test_case_df['Execute'] == 'Yes']

    # Rename columns to avoid name collisions during merging
    filtered_test_case_df = filtered_test_case_df.add_prefix('TestCase_')

    # Initialize a list for storing matched data
    matched_data_list = []

    # Iterate through filtered test cases
    for _, row in filtered_test_case_df.iterrows():
        module_name = row['TestCase_Module']
        try:
            # Load the module-specific sheet
            module_sheet_df = pd.read_excel(data_file, sheet_name=module_name)

            # Rename columns for clarity
            module_sheet_df = module_sheet_df.add_prefix('Module_')

            # Match MTC values between the two DataFrames
            matched_data_df = module_sheet_df[module_sheet_df['Module_MTC'] == row['TestCase_MTC']]

            if not matched_data_df.empty:
                # Merge the data with the filtered test cases
                merged_df = pd.merge(
                    filtered_test_case_df[filtered_test_case_df['TestCase_MTC'] == row['TestCase_MTC']],
                    matched_data_df,
                    left_on='TestCase_MTC',
                    right_on='Module_MTC',
                    how='left'
                )
                matched_data_list.append(merged_df)
        except ValueError:
            print(f"Sheet for module '{module_name}' not found in {data_file}")

    # Combine all matched data into a single DataFrame
    if matched_data_list:
        final_matched_df = pd.concat(matched_data_list, ignore_index=True)
        return final_matched_df
    else:
        print("No matched data found.")
        return pd.DataFrame()

# Example usage
matched_data = get_matched_test_cases(test_case_file, data_file)

if not matched_data.empty:
    print("Matched Data:")
    print(matched_data)
else:
    print("No matched data found.")
