from flask import Flask, request, render_template
from selenium import webdriver # (1) login 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import Select
import io,os,time

from distutils.log import debug
from fileinput import filename
import pandas as pd
#from flask import *
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = os.path.join('staticFiles', 'uploads')
# Define allowed files
ALLOWED_EXTENSIONS = {'csv'}

app = Flask(__name__)
# Configure upload file path flask
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'This is your secret key to utilize session in Flask'

@app.route('/', methods=['GET', 'POST'])
def login_and_upload():
    if request.method == 'POST':
        useremail = request.form.get('useremail')
        password = request.form.get('password')
		
        f = request.files.get('file')
        data_filename = secure_filename(f.filename)
		# Read the filestorage object f as pandas dataframe
        df_uploaded = pd.read_csv(io.StringIO(f.read().decode('utf-8')))
		#df_uploaded_html = df_uploaded.to_html()
				

        driver = webdriver.Chrome()
        driver.get('https://az3.ondemand.esker.com/ondemand/webaccess/asf/home.aspx')
        driver.maximize_window()
        
        driver.find_element(By.XPATH, '//*[@id="ctl03_tbUser"]').send_keys(useremail)
        driver.find_element(By.XPATH, '//*[@id="ctl03_tbPassword"]').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="ctl03_btnSubmitLogin"]').click()
        return df_uploaded.to_html(),data_filename
    return render_template('login.html')  # login.html template

@app.route('/', methods=['GET', 'POST'])
def uploadFile():
	if request.method == 'POST':
	# upload file flask
		f = request.files.get('file')

		# Extracting uploaded file name
		data_filename = secure_filename(f.filename)

		f.save(os.path.join(app.config['UPLOAD_FOLDER'],
							data_filename))

		session['uploaded_data_file_path'] = os.path.join(app.config['UPLOAD_FOLDER'],
					data_filename)

		return render_template('index2.html')
	return render_template("index.html")

@app.route('/show_data')
def showData():
	# Uploaded File Path
	data_file_path = session.get('uploaded_data_file_path', None)
	# read csv
	uploaded_df = pd.read_csv(data_file_path,
							encoding='unicode_escape')
	# Converting to html Table
	uploaded_df_html = uploaded_df.to_html()
	return render_template('show_csv_data.html',
						data_var=uploaded_df_html)

if __name__ == '__main__':
	app.run(debug=True)
