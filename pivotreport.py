import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import PieChart3D, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

# Load the raw data from Excel file into a pandas DataFrame
df = pd.read_excel('/Users/anijahphillip/Downloads/reportdata.xlsx', skiprows=6)

# Get a list of all unique agencies in the data
agencies = df['AGENCY'].unique()

# Create a new workbook
workbook = Workbook()
workbook.remove(workbook.active)


# Loop through each agency and create a worksheet with a pivot table and pie chart
for agency in agencies:
    # Filter the data to include only the current agency
    agency_data = df[df['AGENCY'] == agency]

    # Create a pivot table for the current agency
    pivot_table = pd.pivot_table(agency_data, values='AGENCY', index='02 - Information Security in the Workplace', aggfunc='count', margins=True, margins_name='Grand Total')
    pivot_table.index = pivot_table.index.rename('')
    pivot_table = pivot_table.rename(columns={'AGENCY': agency.upper()})

    worksheet = workbook.create_sheet(title=agency.upper())

    for r in dataframe_to_rows(pivot_table, index=True, header=True):
        worksheet.append(r)

    pie_chart = PieChart3D()
    labels = Reference(worksheet, min_col=1, min_row=3, max_row=5)
    data = Reference(worksheet, min_col=2, max_col=pivot_table.shape[1]+1, min_row=2, max_row=pivot_table.shape[0]+1)
    pie_chart.add_data(data, titles_from_data=True)
    pie_chart.set_categories(labels)
    pie_chart.title = f'{agency.upper()} % Rate'
    worksheet.add_chart(pie_chart, f"C{pivot_table.shape[0]+3}")
    

# Save the workbook to an Excel file
output_file = '/Users/anijahphillip/Downloads/Security_Awarenessreport.xlsx'
workbook.save(output_file)