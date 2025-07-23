from flask import Flask, render_template, request, redirect
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/about')
def about():
    return render_template('About.html')

@app.route('/contact')
def contact():
    return render_template('Contactus.html')

@app.route('/thankyou')
def thank_you():
    return render_template('thankyou.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    email = request.form['email']
    message = request.form['message']

    df = pd.DataFrame([[name, email, message]], columns=["Name", "Email", "Message"])
    
    file_path = 'form_data.xlsx'
    try:
        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            df = pd.concat([existing_df, df], ignore_index=True)
        df.to_excel(file_path, index=False)
    except PermissionError:
        return "Excel file is open. Please close it and try again."

    return redirect('/thankyou')

if __name__ == '__main__':
    app.run(debug=True, port=5001)
