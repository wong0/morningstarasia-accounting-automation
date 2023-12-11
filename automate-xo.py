import os
import subprocess

import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from openpyxl.utils.dataframe import dataframe_to_rows
import configparser

# Create a new configuration object
config = configparser.ConfigParser()

# Set path vars
download_path = '/Users/gw/Downloads/'
date_path = '20230701 20230731'

# Set the values for 'download path' and 'from_to_datestrings'
config['DEFAULT'] = {
    'download_path': download_path,
    'from_to_datestrings': date_path
}

# Save the configuration to a file
with open('automation_configurations_xo.cfg', 'w') as config_file:
    config.write(config_file)

file_path = download_path + 'tasan xo ' + date_path + '.xls'
# subprocess.call(['open', file_path])  # Open the file using the default application

# Load the downloaded Excel file using pandas
# This is a html file masquerading as xls.
df = pd.read_html(file_path)[0]

df.to_excel(download_path+'tasan xo '+ date_path +'.xlsx', index=False)

df = pd.read_excel(download_path + 'tasan xo '+ date_path +'.xlsx')

# Exclude the header and last line with totals
values_df = df.iloc[7:-1, :]

print(values_df.head())

# Load the Excel template using openpyxl
template_file_path = download_path + 'tasan to zoho xo import.xlsx'
template_wb = load_workbook(template_file_path)
template_ws = template_wb.active

# Copy the values to the template starting from the top-left cell
for row in values_df.iterrows():
    template_ws.append(list(row[1].values))

# Save the modified template
output_file_path = download_path + 'tasan to zoho xo import 20230101 20230630.xlsx'
template_wb.save(output_file_path)

subprocess.call(['open', output_file_path])  # Open the file using the default application

# Prompt the user for input to interrupt the script
input("Please make any edits, save on excel and exit. Press Enter to continue...")
