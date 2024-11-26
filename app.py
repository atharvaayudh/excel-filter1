from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename

# Initialize Flask app
app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Home route to upload Excel file.
    """
    if request.method == 'POST':
        # Handle file upload
        file = request.files['file']
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Read Excel file to get column names
            try:
                df = pd.read_excel(filepath)
                columns = df.columns.tolist()
                return render_template('filter.html', columns=columns, file=filename)
            except Exception as e:
                return f"Error processing file: {e}", 400

    return render_template('index.html')


@app.route('/filter', methods=['POST'])
def filter_file():
    """
    Route to filter data by a selected column and generate an Excel file with selected columns.
    """
    try:
        # Get selected column, required columns, and filename from form
        filter_column = request.form['filter_column']
        selected_columns = request.form.getlist('columns')
        filename = request.form['file']
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        # Read the Excel file
        df = pd.read_excel(filepath)
        unique_values = df[filter_column].unique()

        # Prepare output file path
        output_filename = f'filtered_{filename}'
        output_filepath = os.path.join(app.config['RESULT_FOLDER'], output_filename)

        # Create a new Excel file with filtered data
        writer = pd.ExcelWriter(output_filepath, engine='openpyxl')

        for value in unique_values:
            filtered_data = df[df[filter_column] == value]

            # Select only the required columns
            filtered_data = filtered_data[selected_columns]
            
            if not filtered_data.empty:
                sheet_name = str(value)[:31]  # Sheet names must be <= 31 characters
                filtered_data.to_excel(writer, sheet_name=sheet_name, index=False)

        writer.close()  # Finalize the file

        # Send file for download
        return send_file(output_filepath, as_attachment=True)

    except Exception as e:
        return f"Error during processing: {e}", 500


if __name__ == '__main__':
    app.run(debug=True)
