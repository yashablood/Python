from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import os

app = Flask(__name__)

# Configure the SQL database
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data.db'
db = SQLAlchemy(app)

# Define a model for storing Excel data
class ExcelData(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    column1 = db.Column(db.String(50))
    column2 = db.Column(db.String(50))
    # Add more columns as needed

# Ensure the database is created within the application context
@app.before_first_request
def init_db():
    db.create_all()

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    if file:
        df = pd.read_excel(file)

        # Process and store the data in the SQL database
        for index, row in df.iterrows():
            data = ExcelData(column1=row['Column1'], column2=row['Column2'])
            db.session.add(data)
        db.session.commit()

        return jsonify({'message': 'File processed and data saved successfully'}), 200
    else:
        return jsonify({'message': 'No file uploaded'}), 400

@app.route('/data', methods=['GET'])
def get_data():
    data = ExcelData.query.all()
    result = [{'column1': d.column1, 'column2': d.column2} for d in data]
    return jsonify(result), 200

if __name__ == '__main__':
    app.run(debug=True)
