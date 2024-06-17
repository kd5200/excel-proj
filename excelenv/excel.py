import pandas as pd 
import numpy as np 


# Load the Excel files
file_1 = 'initial-file-path'
file_2 = 'file-path-of-dcm-export'



# Map columns to each other so theres no need to rename columns
column_mapping = {

    'Site' : 'Site Name',
    'DCM Placement ID' : 'Placement ID',
    'DCM Placement Name' : 'Placement Name',
    'Placement Start Date' : 'Start Date',
    'Placement End Date' : 'End Date',
    'Size' : 'Dimensions',
    'Ad Server' : 'Ad Type',
    'Click-through URL' : 'Creative Click-Through URL'


}




# Read the sheets into DataFrames
df1 = pd.read_excel(file_1, sheet_name='Sheet1')
df2 = pd.read_excel(file_2, sheet_name='Sheet2')




# Display the initial columns

print('Initial colmns in first sheet:', df1.columns.tolist())
print('Initial colmns in second sheet:', df2.columns.tolist())



# Specify the columns to be deleted from the first sheet
# Set time aside to see which columns aren't needed on the dcm export excel sheet
columns_to_remove = ['Site ID', 'Placement Compatibility', 'Placement Group Type', 'Placement Group ID', 'Primary Placement', 'Payment Source', 'Placement Orientation', 'Placement Duration', 'Placement Publisher Specification', 'Placement Tag Wrapping', 'Placement Tag Wrapping Type', 'Placement Tag Wrapping Measurement Mode', 'Cost Structure', 'Units', 'Rate', 'Cost', 'Opt This Placement Out of Ad Blocking', 'Placement Comments', 'Content Category', 'Placement Strategy', 'Assignment is Active', 'Ad ID', 'Ad Uses Default Landing Page', 'Ad Landing Page ID', 'Ad Landing Page Name', 'Ad Click-Through URL', 'Applied Impression Event Tag IDs']



# Remove the specified columns from the first sheet
df1 = df1.drop(columns=columns_to_remove)

# Display the columns after removal
print('columns after removal in new sheet:', df1.columns.tolist())

# Map columns to match naming convention 
df2_mapped = df2.rename(columns=column_mapping)



# Identify common columns (why do I need this?)
common_columns = df1.columns.intersection(df2.columns)


# Filter the DataFrames to include only the common columns
df1_filtered = df1[common_columns]
df2_filtered = df2[common_columns]


# Display the common columns
print('Common columns:', common_columns.to_list())


# Compare the filtered DataFrames and highlight differences
comparison_values = df1_filtered.values == df2_filtered.values



# Create a DataFrame to store differences
df_differences = pd.DataFrame(np.where(comparison_values, np.nan, 'Difference'), index=df1_filtered.index, columns=df1_filtered.columns)


# Save the differences to a new Excel file
with pd.ExcelWriter('QA.xlsx') as writer:
    df1_filtered.to_excel(writer, sheet_name='Sheet1_filtered', index=False)
    df2_filtered.to_excel(writer, sheet_name='Sheet2_filtered', index=False)
    df_differences.to_excel(writer, sheet_name='Sheet_QA', index=False)

print("Differences have een highlighted in the 'QA.xlsx' sheet")
