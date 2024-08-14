from flask import Flask, request, jsonify
import pandas as pd
import os
from fpdf import FPDF
import matplotlib.pyplot as plt


global_arr=[]
app=Flask(__name__)
def generate_pdf(data):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, "Analysis Results", ln=True, align='C')

    if isinstance(data, list):
        for idx, value in enumerate(data):
            if isinstance(value, (int, float)):
                pdf.cell(200, 10, f"Sheet {idx+1} Result: {value}", ln=True)
            else:
                pdf.cell(200, 10, "Invalid data type for result", ln=True)
    else:
        pdf.cell(200, 10, "Invalid data format", ln=True)

    pdf.output("analysis_results.pdf")

@app.route('/upload',methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file=request.files['file']
    if file.filename=='':
        return 'No selected file'
    if file:
        file.save('uploads/'+file.filename)
        base_url=os.path
        file_url=base_url.abspath('uploads/'+file.filename)
        excel_data=pd.ExcelFile('uploads/'+file.filename)
        num_sheets=len(excel_data.sheet_names)
        return f'File uploaded successfully. Full URL: {file_url}, Number of sheets: {num_sheets}'
@app.route('/report', methods=['POST'])
def report():
    data=request.get_json()
    url=data.get('url')
    sheets=data.get('sheets')
    arr=[]
    for sheet in sheets:
        df = pd.read_excel(url,sheet_name=sheet['name'])
        column_sum=0
        for column in sheet['columns']:
            if sheet['action']=='avg':
                column_sum+=df[column].mean()
            if sheet['action'] == 'sum':
                column_sum += df[column].sum()
        arr.append(column_sum)
    global global_arr
    global_arr=arr
    generate_graphs(url, sheets)  # Generate graphs
    return jsonify(arr)


def generate_graphs(excel_data, sheets):
    for sheet in sheets:
        df = pd.read_excel(excel_data, sheet_name=sheet['name'])

        plt.figure(figsize=(6, 4))
        for column in sheet['columns']:
            plt.plot(df[column], label=column)
        plt.xlabel('Rows')
        plt.ylabel('Values')
        plt.title(f'Graph for {sheet["name"]}')
        plt.legend()

        graph_filename = f"{sheet['name']}_graph.png"
        plt.savefig(graph_filename)
        plt.close()
if __name__ =='__main__':
    app.run()
    generate_pdf(global_arr)

