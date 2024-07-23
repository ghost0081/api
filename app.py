import os
from flask import Flask, request, render_template, send_from_directory
import openpyxl

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'jpg', 'jpeg', 'png', 'gif'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit_form():
    # Collect form data
    formData = {
        'apMainCode': request.form['apMainCode'],
        'apName': request.form['apName'],
        'date': request.form['date'],
        'location': request.form['location'],
        'visitSpoc': request.form['visitSpoc'],
        'modeOfAudit': request.form['modeOfAudit'],
        'noOfEmployees': request.form['noOfEmployees'],
        'noOfClientMappedToAp': request.form['noOfClientMappedToAp']
    }

    # Process user table data
    userTableData = []
    for i in range(len(request.form.getlist('userId[]'))):
        userId = request.form.getlist('userId[]')[i]
        userName = request.form.getlist('userName[]')[i]
        segment = request.form.getlist('segment[]')[i]
        approvedAddress = request.form.getlist('approvedAddress[]')[i]
        actualLocatedAt = request.form.getlist('actualLocatedAt[]')[i]
        remarks = request.form.getlist('remarks[]')[i]
        authorizedPersonReply = request.form.getlist('authorizedPersonReply[]')[i]
        
        evidence_file = request.files.getlist('evidence[]')[i] if 'evidence[]' in request.files else None
        if evidence_file and allowed_file(evidence_file.filename):
            filename = os.path.join(app.config['UPLOAD_FOLDER'], f"user_{i+1}_{evidence_file.filename}")
            evidence_file.save(filename)
        else:
            filename = ''

        userTableData.append((userId, userName, segment, approvedAddress, actualLocatedAt, remarks, authorizedPersonReply, filename))

    # Process List Of Boards data
    boardList = [
        "Main display board with words franchisee",
        "Notice board of the trading Member containing all details / information prescribed from time to time, are displayed at the AP location",
        "SEBI registration certificate of the trading Member and registration letter issued by the Exchange is displayed at location",
        "Message of investor alert",
        "AP registration certificate - BSE / NSE / MCX / NCDEX",
        "Message for investor (Do's & Don'ts)",
        "Message for investor Rights & Responsibilities",
        "Investor Charter",
        "Grievance Redressal Mechanism (SEBI circular CIR/MIRSD/3/2014 dated August 28, 2014)"
    ]

    boardTableData = []
    for i in range(len(request.form.getlist('yesNo[]'))):
        srNo = i + 1
        listOfBoards = boardList[i]
        yesNo = request.form.getlist('yesNo[]')[i]
        boardRemarks = request.form.getlist('boardRemarks[]')[i]
        boardAuthorizedPersonReply = request.form.getlist('boardAuthorizedPersonReply[]')[i]
        
        board_evidence_file = request.files.getlist('boardEvidence[]')[i] if 'boardEvidence[]' in request.files else None
        if board_evidence_file and allowed_file(board_evidence_file.filename):
            filename = os.path.join(app.config['UPLOAD_FOLDER'], f"board_{i+1}_{board_evidence_file.filename}")
            board_evidence_file.save(filename)
        else:
            filename = ''

        boardTableData.append((srNo, listOfBoards, yesNo, boardRemarks, boardAuthorizedPersonReply, filename))

    # Create Excel file
    wb = openpyxl.Workbook()
    ws = wb.active

    # Save form data
    ws.append(["AP Main Code", "AP Name", "Date", "Location", "Visit SPOC", "Mode Of Audit", "No Of Employees", "No Of Client Mapped To AP"])
    ws.append([formData['apMainCode'], formData['apName'], formData['date'], formData['location'], formData['visitSpoc'], formData['modeOfAudit'], formData['noOfEmployees'], formData['noOfClientMappedToAp']])

    # Save user data
    ws.append([])
    ws.append(["User ID", "User Name", "Segment", "Approved Address", "Actual Located At", "Remarks", "Authorized Person Reply", "Evidence"])
    for row in userTableData:
        ws.append(row)

    # Save List Of Boards data
    ws.append([])
    ws.append(["Sr No.", "List Of Boards", "Yes/No", "Remarks", "Authorized Person Reply", "Evidence"])
    for row in boardTableData:
        ws.append(row)

    file_path = os.path.join('static', 'result.xlsx')
    wb.save(file_path)

    return f"Form submitted successfully! <a href='/static/result.xlsx'>Download the Excel file</a>"

if __name__ == '__main__':
    # Ensure the uploads folder exists
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    
    app.run(debug=True)
