from flask import Flask, render_template
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route('/')
def display_workbook():
    # Load the workbook
    workbook = load_workbook('/Users/anijahphillip/Downloads/Security_Awarenessreport.xlsx')

    # Extract the sheet names
    sheet_names = workbook.sheetnames

    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Render the template and pass the sheet names and workbook data
    return render_template('workbook.html', sheet_names=sheet_names, workbook=workbook)

if __name__ == '__main__':
    app.run(debug=True)
