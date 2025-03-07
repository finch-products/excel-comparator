from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file1' not in request.files or 'file2' not in request.files:
        return "Please upload both Excel files!"
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    file1_path = os.path.join(UPLOAD_FOLDER, file1.filename)
    file2_path = os.path.join(UPLOAD_FOLDER, file2.filename)
    file1.save(file1_path)
    file2.save(file2_path)
    
    return render_template('compare.html', file1=file1.filename, file2=file2.filename)

@app.route('/compare', methods=['POST'])
def compare():
    file1_path = os.path.join(UPLOAD_FOLDER, request.form['file1'])
    file2_path = os.path.join(UPLOAD_FOLDER, request.form['file2'])
    
    xls1 = pd.ExcelFile(file1_path)
    xls2 = pd.ExcelFile(file2_path)
    sheets = xls1.sheet_names
    today_date = datetime.today().strftime('%Y-%m-%d')
    output_file = os.path.join(RESULT_FOLDER, f"Comparison_{today_date}.xlsx")
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for sheet in sheets:
            df1 = pd.read_excel(xls1, sheet_name=sheet)
            df2 = pd.read_excel(xls2, sheet_name=sheet)
    
            diff_df = df1.merge(df2, how='outer', indicator=True).query('_merge == "left_only"').drop('_merge', axis=1)
            
            if not diff_df.empty:
                diff_df.to_excel(writer, sheet_name=f"{sheet}_{today_date}", index=False)
                print(f"Differences found in {sheet} and stored in {output_file} under sheet {sheet}_{today_date}")
            else:
                print(f"No differences found in {sheet}")
    
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
