# views.py
import sys
from flask import render_template
from flask import request
from reader import getAreas
from app import app

@app.route('/')
def index():
   #return("Holi :3")
   return render_template("fileUpload.html")

@app.route('/success', methods = ['GET', 'POST'])  
def success():  
   if request.method == 'POST':
      print('ying', file=sys.stdout)
      f = request.files['file']
      f.save('./uploads/excelFile.xlsx')
      areas=getAreas('./uploads/excelFile.xlsx')
      return render_template("report.html", areas=areas) 
   return render_template("uploadError.html") 
 
@app.route('/about')
def about():
   return render_template("about.html")
