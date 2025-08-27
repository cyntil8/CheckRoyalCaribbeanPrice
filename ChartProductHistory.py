import openpyxl
from openpyxl.chart import LineChart, Reference, BarChart
from openpyxl.utils import get_column_letter
import argparse
from datetime import datetime

def build_chart_from_description(filename, description):
    wb = openpyxl.load_workbook(filename)

    results = []
    
    for sheetname in wb.sheetnames:
        if "Chart - " in sheetname or sheetname == "Sheet":
            continue
        sheet = wb[sheetname]
        date = datetime.strptime(sheetname, "%Y-%m-%d (%H %M)")
        for row in sheet.iter_rows(min_row=2, values_only=True):
            desc = row[1]
            amount = row[3]
            if desc and description.lower() in str(desc).lower():
                results.append((date, amount))

    if results:

        results.sort(key=lambda x: x[0])

        chart_sheet_name = "Chart - " + description
        if chart_sheet_name in wb.sheetnames:
            chart_sheet = wb[chart_sheet_name]
            wb.remove(chart_sheet)
        chart_sheet = wb.create_sheet(chart_sheet_name, 0)

        chart_sheet.append(["Date", "Amount"])
        for date, amount in results:
            chart_sheet.append([date, amount])

        chart = LineChart()
        chart.title = f"Results for '{description}'"
        chart.x_axis.title = "Date"
        chart.y_axis.title = "Amount"

        data = Reference(chart_sheet, min_col=2, min_row=1, max_row=len(results)+1)
        cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=len(results)+1)

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        chart_sheet.add_chart(chart, "E5")

        wb.save(filename)

parser = argparse.ArgumentParser()
parser.add_argument("-f", "--filename", help="filename") 
parser.add_argument("-p", "--product", help="Product") 
args = parser.parse_args()

build_chart_from_description(args.filename, args.product)
