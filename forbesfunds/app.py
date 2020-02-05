#!/usr/bin/python
import pandas as pd
import numpy as np
from flask import Flask,request, render_template, make_response
import os
import urllib.request
#UPLOAD_FOLDER = '/'

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/getcsv', methods = ['POST'])
def gecsv():
	if request.method == 'POST':
		f = request.files['file']
		f.save(f.filename)
		import pandas as pd
		wb = pd.ExcelFile(f.filename)
		template=pd.read_excel(f.filename,wb.sheet_names[0])
		needs=pd.read_excel(f.filename,wb.sheet_names[3])
		data=pd.read_excel(f.filename,wb.sheet_names[4])
		allorgs=pd.read_excel(f.filename,wb.sheet_names[1])
		f.close()        
		allorgs.rename(columns={"Organization ": "Name of organization"}, inplace=True)
		all_orgs=pd.DataFrame(allorgs['Name of organization'])
		finalpd=all_orgs.merge(needs, on='Name of organization', how='left')
		finalpd[['Date organization completed Needs Assessment','Completed Needs Assessment (Y/N)']]=finalpd['Timestamp'].astype('str').str.split(" ", expand=True)
		finalpd['Completed Needs Assessment (Y/N)'] = np.where(pd.isnull(finalpd['Completed Needs Assessment (Y/N)']), 'no','yes')
		finalpd["Targeted Demographic"]= needs["Who is your organization's targeted demographic?"].astype(str) + ", "+ needs["Who is your organization's targeted demographic?.1"].astype(str)+ ", " + needs["Who is your organization's targeted demographic?.2"].astype(str)
		finalpd=finalpd.drop(['Timestamp','Website','Mission of the Organization',"Who is your organization's targeted demographic?","Who is your organization's targeted demographic?.1","Who is your organization's targeted demographic?.2",'Vision of the Organization','Overarching Goals of the Organization','Overarching Strategy of the Organization',                      'Shared Beliefs & Values within the Organization','Board Composition & Commitment','Board Governance','Board Involvement & Support','CEO/ED Experience & Standing','CEO/ED Organizational Leadership/Effectiveness','CEO/ED Analytical & Strategic Thinking','CEO/ED Financial Judgment','Board & CEO/ED Appreciation of Power Issues',                      'Community Presence & Standing of the Organization',"Organization's Ability to Motivate & Mobilize Constituents",'Strategic Planning','Evaluation/Performance Measurement','Evaluation & Organizational Learning','Use of Research Data to Support Program Planning & Advocacy','Program Growth & Replication','New Program Development',                      'Monitoring of Program Landscape','Assessment of External Environment & Community Needs','Influencing of Policy-making','Partnerships & Alliances','Organizing','Senior Management Team','Staff','Dependence of Management Team & Staff on CEO/ED','Shared References & Practices','Goals/ Performance Targets','Program Relevance & Integration','Funding Model','Fund Development Planning','Financial Planning/ Budgeting','Financial Operations Management','Organizational Processes','Decision Making Processes','Knowledge Management','lnterfunctional Coordination & Communication','Human Resources Planning','Staffing Levels','Recruiting, Development, & Retention of Management','Recruiting, Development, & Retention of General Staff','Volunteer Management','Constituent Involvement','Operational Planning','Skills, Abilities, & Commitment of Volunteers','Fundraising','Board Involvement & Participation in Fundraising','Revenue Generation','Communications Strategy','Communications & Outreach','Telephone & Fax','Computers, Applications, Network, & Email','Databases/ Management Reporting Systems','Buildings & Office Space','Management of Legal & Liability Matters'], axis=1)
		finalpd.rename(columns={'Website address (put N/A if you do not have a website)': "Website",
                        'What is the budget size of your organization':"Organization's Budget Size",
                        "What is your organization's mission statement?":"Organization's Mission Statement",
                        "Which system is your organization a part of serving?":"System Served",
                       'Is your organization currently (or formerly)... [a Greater Pittsburgh Nonprofit Partnership (GPNP) member]':"GPNP Member",
                       'Is your organization currently (or formerly)... [receiving/received Executive in Residence (EIR) assistance (individual or cohort)]':"EIR Recipient",
                       'Is your organization currently (or formerly)... [a Management Assistance Grant recipient]':"MAG recipient",
                       'Is your organization currently (or formerly)... [an UpPrize recipient]':"UpPrize Recipient",
                       'Is your organization currently (or formerly)... [a non-EIR cohort attendee]':"Cohort Attendee"}, inplace=True)

		temp=data[['RecipientName','RecipientAddressBlock','RecipientCity','RecipientState','RecipientZip','RecipientFullAddress']]
		address=temp.drop_duplicates(subset='RecipientName', keep="first")
		address.rename(columns={'RecipientName':'Name of organization','RecipientAddressBlock':'Street Address','RecipientCity':'City','RecipientState':'State','RecipientZip':'Zip','RecipientFullAddress':'FullAddress'}, inplace=True)
		org_and_address=finalpd.merge(address, on='Name of organization', how='left')		
		os.system(f"rm -f '{f.filename}'")
		resp = make_response(org_and_address.to_csv())
		resp.headers["Content-Disposition"] = "attachment; filename=Insightly.csv"
		resp.headers["Content-Type"] = "text/csv"
		return resp

 
if __name__ == '__main__':
   app.run(debug=True)
