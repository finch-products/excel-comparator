from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import shutil

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def delete_files():
    """Delete all files in the upload and result folders."""
    for folder in [UPLOAD_FOLDER, RESULT_FOLDER]:
        if not os.path.exists(folder):
            print(f"Folder {folder} does not exist. Skipping deletion.")
            continue

        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                    print(f"Deleted file: {file_path}")
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    print(f"Deleted directory: {file_path}")
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

@app.route('/')
def index():
    # Delete all files when the index page is loaded
    delete_files()
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
    output_file = os.path.join(RESULT_FOLDER, "Comparison.xlsx")

    differences = []

    for sheet in sheets:
        df1 = pd.read_excel(xls1, sheet_name=sheet)
        df2 = pd.read_excel(xls2, sheet_name=sheet)

        # Compare both ways to find all differences
        diff_df1 = df1.merge(df2, how='outer', indicator=True).query('_merge == "left_only"').drop('_merge', axis=1)
        diff_df2 = df2.merge(df1, how='outer', indicator=True).query('_merge == "left_only"').drop('_merge', axis=1)

        if not diff_df1.empty or not diff_df2.empty:
            diff_df1.insert(0, "Sheet Name", sheet)
            diff_df2.insert(0, "Sheet Name", sheet)
            differences.append(diff_df1)
            differences.append(diff_df2)

    if differences:
        final_df = pd.concat(differences, ignore_index=True)
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name="Differences", index=False)
        
        # Send the file and then delete it
        response = send_file(output_file, as_attachment=True)
        delete_files()  # Delete files after sending the response
        return response
    else:
        delete_files()  # Delete files if no differences are found
        return "No differences found between the two files."

if __name__ == '__main__':
    app.run(debug=True)