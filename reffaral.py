import streamlit as st
import pandas as pd
import io

# Define the list of departments
departments = [
    "BIOCHEMISTRY", "CLINICAL PATHOLOGY", "CYTO-PATHOLOGY", "HAEMATOLOGY", 
    "IMMUNOLOGY", "MICROBIOLOGY", "SEROLOGY", "HISTO-PATHOLOGY", 
    "IMMUNOHISTOCHEMISTRY", "MOLECULAR BIOLOGY", "Flowcytometry", 
    "PHARMACOGENETICS LAB", "PCR LAB", "ECG", "CARDIAC TEST", "X-RAY", "USG 3D/4D"
]

# Load Excel data
@st.cache_data
def load_sheets(file):
    xl = pd.ExcelFile(file)
    return xl.sheet_names

@st.cache_data
def load_data(file, sheet_name):
    return pd.read_excel(file, sheet_name=sheet_name)

# Function to add a grand total row to a DataFrame
def add_grand_total(df):
    total_row = df.sum(numeric_only=True, skipna=True)
    #total_row['Mkt Code'] = 'Grand Total'
    df_with_total = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df_with_total, total_row

# Streamlit UI
st.title("DMFR Referral Sheet")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    sheet_names = load_sheets(uploaded_file)
    
    main_sheet_name = st.selectbox("Select the main sheet (Test Wise):", sheet_names)
    referral_policy_name = st.selectbox("Select the referral policy sheet:", sheet_names)
    doctor_wise_name = st.selectbox("Select the doctor wise sheet:", sheet_names)

    if main_sheet_name and referral_policy_name:
        main_sheet = load_data(uploaded_file, main_sheet_name)
        referral_policy = load_data(uploaded_file, referral_policy_name)
        doctor_wise_sheet = load_data(uploaded_file, doctor_wise_name)
        
        if st.button("Process Data"):
            try:

                # Initialize the 'Ref Amount' column with empty strings
                main_sheet['Ref Amount'] = ''

                # Replace 'Test Name' and 'Department' with the actual column names in your Excel sheet
                test_name_col = 'Referral Doctor'  # Change this to the actual column name for test names
                department_col = 'Department'  # Change this to the actual column name for departments
                total_sale_col = 'Total Sale'
                invoice_no_col = 'Invoice No'
                invoice_id_col = 'Invoice Id'


                # Check if columns exist
                if test_name_col not in main_sheet.columns or department_col not in main_sheet.columns or total_sale_col not in main_sheet.columns or invoice_no_col not in main_sheet.columns or invoice_id_col not in doctor_wise_sheet.columns:
                    st.error("Columns for test name, department, total sale, invoice number, or invoice ID not found. Please check your Excel file.")
                else:
                    # Initialize the 'Ref Amount' column with NaN
                    main_sheet['Ref Amount'] = pd.NA
                    # Iterate through the rows in the main sheet
                    for idx, row in main_sheet.iterrows():
                        test_name = row[test_name_col]
                        department = row[department_col]

                        if department in departments:
                            # Find the column index for the department in the 'REFERRAL POLICY' sheet
                            column_index = departments.index(department) + 3  # +2 because index starts from 0 and columns start from 1
                           
                            # Perform the lookup
                            matched_value = referral_policy.loc[referral_policy['DOCTOR NAME'] == test_name, referral_policy.columns[column_index]]

                            if not matched_value.empty:
                                main_sheet.at[idx, 'Ref Amount'] = matched_value.values[0]
                    # Convert 'Ref Amount' column to numeric, coercing errors to NaN
                    main_sheet['Ref Amount'] = pd.to_numeric(main_sheet['Ref Amount'], errors='coerce')
                     # Fill NaN values in 'Ref Amount' with 0
                    main_sheet['Ref Amount'].fillna(0, inplace=True)

                    # Calculate the 'Referral' column
                    main_sheet['Referral'] = main_sheet['Total Sale'] * main_sheet['Ref Amount']
                     # Fill NaN values in 'Referral' with 0 (in case there are any)
                    main_sheet['Referral'].fillna(0, inplace=True)
                            
                # Create a pivot table
                    pivot_table = main_sheet.pivot_table(
                        index=['Invoice No'], 
                        values=['Referral'], 
                        aggfunc='sum'
                    )
                # Merge pivot table with the doctor wise sheet on Invoice Id
                    doctor_wise_sheet = doctor_wise_sheet.merge(pivot_table, left_on=invoice_id_col, right_index=True, how='left')

                    # Fill NaN values in 'Referral' with 0 (in case there are any) after the merge
                    doctor_wise_sheet['Referral'].fillna(0, inplace=True)

                    # Rename the merged column to 'CDR'
                    doctor_wise_sheet.rename(columns={'Referral': 'CDR'}, inplace=True)

                    # Calculate the 'CDR Percent' column
                    doctor_wise_sheet['CDR Percent'] = (
                        (doctor_wise_sheet['CDR'] / doctor_wise_sheet['Actual Total Sale'])
                    )

                    # Calculate the 'Actual Total Discount' column
                    doctor_wise_sheet['Actual Referral'] = (
                        doctor_wise_sheet['CDR'] - 
                        doctor_wise_sheet['Actual Total Discount']
                    )
                    # Ensure 'Actual Referral' is non-negative
                    doctor_wise_sheet['Actual Referral'] = doctor_wise_sheet['Actual Referral'].clip(lower=0)
                    # Calculate the 'Ref Percent' column
                    doctor_wise_sheet['Ref Percent'] = (
                        doctor_wise_sheet['Actual Referral']/ 
                        doctor_wise_sheet['Actual Total Sale']
                    )

                     # Calculate the 'Actual Net Sale' column
                    doctor_wise_sheet['Actual Net Sale'] = (
                        doctor_wise_sheet['Actual Total Sale'] - 
                        doctor_wise_sheet['Actual Total Discount'] - 
                        doctor_wise_sheet['Actual Referral']
                    )
                     # Calculate the 'Discount Percent' column
                    doctor_wise_sheet['Discount Percent'] = (
                        (doctor_wise_sheet['Actual Total Discount'] / doctor_wise_sheet['Actual Total Sale'])
                    )
                    # Ensure 'Discount Percent' is beside 'Actual Total Discount'
                    columns = list(doctor_wise_sheet.columns)
                    discount_index = columns.index('Actual Total Discount') + 1
                    columns.insert(discount_index, columns.pop(columns.index('Discount Percent')))
                    doctor_wise_sheet = doctor_wise_sheet[columns]

                    # Create a BytesIO object to save the Excel file
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        main_sheet.to_excel(writer, sheet_name=main_sheet_name, index=False)
                        referral_policy.to_excel(writer, sheet_name=referral_policy_name, index=False)
                        pivot_table.to_excel(writer, sheet_name='Pivot Data')
                        doctor_wise_sheet.to_excel(writer, sheet_name=doctor_wise_name, index=False)

                        # Write the "Doctor Wise" sheet separated by "Mkt Code"
                        for mkt_code in doctor_wise_sheet['Mkt Code'].unique():
                            if pd.isna(mkt_code):
                                mkt_code_df = doctor_wise_sheet[doctor_wise_sheet['Mkt Code'].isna()]
                                sheet_name = 'Walking Patient'
                            else:
                                mkt_code_df = doctor_wise_sheet[doctor_wise_sheet['Mkt Code'] == mkt_code]
                                sheet_name = f'{mkt_code}'
                            
                            # Separate rows where 'Invoice Id' starts with "B2"
                            b2_df = mkt_code_df[mkt_code_df[invoice_id_col].str.startswith('B2', na=False)]
                            other_df = mkt_code_df[~mkt_code_df[invoice_id_col].str.startswith('B2', na=False)]

                            # Add grand total rows
                            b2_df_with_total, b2_total_row = add_grand_total(b2_df)
                            other_df_with_total, other_total_row = add_grand_total(other_df)

                            # Combine both DataFrames, inserting blank rows and a header for the second part
                            combined_df = pd.concat([b2_df_with_total, pd.DataFrame(columns=mkt_code_df.columns, index=range(3)), other_df_with_total], ignore_index=True)

                            # Write the DataFrame to the Excel file
                            combined_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

                            # Get the workbook and worksheet objects
                            workbook  = writer.book
                            worksheet = writer.sheets[sheet_name]

                            # Define the format for the second header
                            header_format = workbook.add_format({
                                'bold': True,
                                'text_wrap': True,
                                'valign': 'top',
                                #'fg_color': '#D7E4BC',
                                'border': 1})

                            # Write the header for the second part
                            for col_num, value in enumerate(mkt_code_df.columns.values):
                                worksheet.write(len(b2_df_with_total) + 3, col_num, value, header_format)

                            overall_total = b2_total_row.add(other_total_row, fill_value=0)
                            worksheet.write(len(combined_df) + 2, 4, 'Grand Total')
                            for col_num, value in enumerate(overall_total):
                                worksheet.write(len(combined_df) + 2, col_num + 5, value)
                            

                    output.seek(0)

                    st.download_button(
                        label="Download updated Excel file",
                        data=output,
                        file_name="updated_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error: {str(e)}")
else:
    st.info("Please upload an Excel file to get started.")