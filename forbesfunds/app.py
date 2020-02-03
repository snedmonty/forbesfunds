#!/usr/bin/python
import pandas as pd
from flask import Flask,request, render_template, make_response
#from werkzeug import secure_filename
import os
#import magic
import urllib.request
#UPLOAD_FOLDER = '/'

app = Flask(__name__)
#app.secret_key = "secret key"
#app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
#app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/getcsv', methods = ['POST'])
def gecsv():
	if request.method == 'POST':
		f = request.files['file']
		f.save(f.filename)
		import pandas as pd
		wb = pd.ExcelFile('Insightly Template and Data_12042019.xlsx')

		template=pd.read_excel('Insightly Template and Data_12042019.xlsx',wb.sheet_names[0])
		needs=pd.read_excel('Insightly Template and Data_12042019.xlsx',wb.sheet_names[3])
		data=pd.read_excel('Insightly Template and Data_12042019.xlsx',wb.sheet_names[4])
		allorgs=pd.read_excel('Insightly Template and Data_12042019.xlsx',wb.sheet_names[1])['Organization ']
#allorgs=allorgs.str.lower()

		temp_pd=needs[['Name of organization', 
               'Timestamp',
               'Number of full-time staff',
               'Number of part-time staff', 
               'Number of volunteers',
               'Number of board members']]

		newdata = data.drop_duplicates(subset='RecipientName', keep="first")
		newone=pd.merge( temp_pd,newdata, right_on='RecipientName', left_on='Name of organization')
		#newone.to_csv('Insightly.csv', encoding='utf-8', index=False)
		resp = make_response(newone.to_csv())
		resp.headers["Content-Disposition"] = "attachment; filename=Insightly.csv"
		resp.headers["Content-Type"] = "text/csv"
		return resp



if __name__ == '__main__':
   app.run(debug=True)
