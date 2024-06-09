import pandas as pd 
import numpy as np 


# Load the Excel files
file_1 = 'initial-file-path'
file_2 = 'file-path-of-dcm-export'





# Read the sheets into DataFrames
df1 = pd.read_excel(file_1, sheet_name='Sheet1')
df2 = pd.read_excel(file_2, sheet_name='Sheet2')



# Display the initial columns

print('Initial colmns in first sheet:', df1.columns.tolist())
print('Initial colmns in second sheet:', df2.columns.tolist())



# Specify the columns to be deleted from the first sheet




# Remove the specified columns from the first sheet



# Display the columns after removal



# Identify common columns




# Filter the DataFrames to include only the common columns



# Display the common columns



# Compare the filtered DataFrames and highlight differences




# Create a DataFrame to store differences



# Save the differences to a new Excel file

