import pandas as pd
from datetime import datetime
# File paths (hardcoded for now)
test_case_file = r"C:\Users\Tmp Admin\PycharmProjects\SeleniumPython\PythonProject\utilities\ALM_Data.xlsx"
second_file = r"C:\Users\Tmp Admin\PycharmProjects\SeleniumPython\PythonProject\utilities\Test_Case.xlsx"

def get_matched_test_cases():
    """
    Fetches test cases from two Excel files where 'Execute' is marked 'Yes'
    and matches data based on 'MTC' and 'Module' columns.

    Returns:
        pd.DataFrame: Final dataframe containing all matched test case data.
    """
    # Load test cases where 'Execute' is 'Yes'
    test_case_df = pd.read_excel(test_case_file)
    filtered_test_case_df = test_case_df[test_case_df['Execute'] == 'Yes']

    # Initialize a list for storing matched data
    matched_data_list = []

    # Iterate through filtered test cases
    for _, row in filtered_test_case_df.iterrows():
        module_name = row['Module']
        try:
            # Load the module-specific sheet
            module_sheet_df = pd.read_excel(second_file, sheet_name=module_name)
            matched_data_df = module_sheet_df[module_sheet_df['MTC'] == row['MTC']]

            if not matched_data_df.empty:
                merged_df = pd.merge(
                    filtered_test_case_df[filtered_test_case_df['MTC'] == row['MTC']],
                    matched_data_df,
                    on='MTC',
                    how='left'
                )
                matched_data_list.append(merged_df)
        except ValueError:
            print(f"Sheet for module '{module_name}' not found in {second_file}")

    # Combine all matched data into a single DataFrame
    if matched_data_list:
        final_matched_df = pd.concat(matched_data_list, ignore_index=True)
        return final_matched_df
    else:
        print("No matched data found.")
        return pd.DataFrame()

