from flask import Flask, render_template, request, send_file, url_for
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import (
    PieChart3D,
    Reference
)
import os 

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_report', methods=['POST'])
def generate_report():
    if request.method == 'POST':
        # Get the uploaded file
        uploaded_file = request.files['file']

        # Load the raw data from the Excel file into a pandas DataFrame
        df = pd.read_excel(uploaded_file, skiprows=6)

        # Get a list of all unique agencies in the data
        agencies = df['AGENCY'].unique()

        # Create a new workbook
        workbook = Workbook()
        workbook.remove(workbook.active) 

        worksheet_name = ""

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
        filename = 'Security_Awareness_generatedreport.xlsx'
        workbook.save(filename)

        # Set the download link to the generated report
        download_link = url_for('download_report', filename=filename)

        return render_template('index.html', download_link=download_link, filename=filename)

    return render_template('index.html')


@app.route('/download_report/<filename>')
def download_report(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
