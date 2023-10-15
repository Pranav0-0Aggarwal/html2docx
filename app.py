import os
from flask import Flask, render_template, request, redirect, url_for, send_file
from docx import Document
from htmldocx2 import HtmlToDocx
import requests
import logging

app = Flask(__name__)

script_dir = os.path.dirname(os.path.realpath(__file__))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/completed')
def completed():
    return render_template('completed.html')

def convert_html_to_docx(html_content):
    doc = Document()
    html_parser = HtmlToDocx()
    html_parser.parse_html_string(html_content)
    return doc

@app.route('/convert', methods=['POST'])
def convert():
    html_content = request.form['html_content']
    doc = convert_html_to_docx(html_content)
    output_file = os.path.join(script_dir, 'static', 'output.docx')
    return redirect(url_for('completed'))

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_file = request.files['file']
    if uploaded_file:
        html_content = uploaded_file.read().decode('utf-8')
        doc = convert_html_to_docx(html_content)
        output_file = os.path.join(script_dir, 'static', 'output.docx')
    return redirect(url_for('completed'))

@app.route('/convert_url', methods=['POST'])
def convert_url():
    try:
        url = request.form['url']
        response = requests.get(url)
        if response.status_code == 200:
            html_content = response.text
            doc = convert_html_to_docx(html_content)
            output_file = os.path.join(script_dir, 'static', 'output.docx')
            return redirect(url_for('download'))
        else:
            return 'Error: Could not retrieve URL content'
    except Exception as e:
        logging.exception(f'Error converting URL: {e}')
        return 'Error: Could not convert URL'

@app.route('/download')
def download():
    try:
        output_file = os.path.join(script_dir, 'static', 'output.docx')
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        logging.exception(f'Error downloading file: {e}')
        return 'Error: Could not download file'

if __name__ == '__main__':
    app.run(debug=True)
