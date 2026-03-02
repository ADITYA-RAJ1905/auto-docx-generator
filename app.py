from flask import Flask, render_template, request, redirect, url_for, send_file, session
from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import sessionmaker, declarative_base
import os
import uuid
# Import logic modules
import pandas as pd
from logic.process_doc1 import process_doc1
from logic.process_doc2 import process_doc2
from logic.process_doc3 import process_doc3
from logic.process_doc4 import process_doc4
from datetime import datetime
from sqlalchemy.orm import declarative_base
Base = declarative_base()

app = Flask(__name__)
app.secret_key = "super-secret-key"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# SQLAlchemy Setup
engine = create_engine("mysql+pymysql://root:Vidyad%401905@localhost/proposal")
Session = sessionmaker(bind=engine)
db_session = Session()
Base = declarative_base()

# UserDocs table
class UserDocs(Base):
    __tablename__ = "userdocs"
    id = Column(Integer, primary_key=True)
    OEM_details = Column(String(10))
    OEM_authorization = Column(String(10))
    GAR = Column(String(10))
    BQ = Column(String(10))
    case_id=Column(String(50))

# Import these models into logic files too
class RecordDoc1(Base):
    __tablename__ = "records_doc1"
    case_id=Column(String(255))
    id = Column(Integer, primary_key=True)
    software_module = Column(String(255))
    address = Column(String(255))
    l_valid=Column(Integer)
    installation=Column(Integer)
    certificate_days=Column(Integer)
    

class CaseDetailsDoc1(Base):
    __tablename__ = "SOW_primary"
    id = Column(Integer, primary_key=True)
    case_id = Column(String(50))  # could also be ForeignKey('records_doc1.case_id')
    desc1 = Column(String(255))
    desc2 = Column(String(255))
    desc3 = Column(String(255))
    L1 = Column(String(255))
    L2 = Column(String(255))
    L3 = Column(String(255))
    
class RecordDoc2(Base):
    __tablename__ = "records"
    case_id=Column(String(255))
    id = Column(Integer, primary_key=True)
    tag = Column(String(255))
    subject = Column(String(255))
    file_no = Column(String(255))
    material = Column(String(255))
    vendor = Column(String(255))
    tender_type = Column(String(255))
    user_disha_file = Column(String(255))
    BQ1_date = Column(String(255))
    Proposal_curr = Column(String(255))
    BQ1_price = Column(Integer)
    BQ2_price = Column(Integer)
    LPR_PO = Column(Integer)
    BQ_per = Column(Integer)
    FY = Column(String(255))
    CURR_EXC_RATE = Column(Integer)
    LPR_UNIT_PRICE = Column(Integer)
    PRICE_DIFF = Column(Integer)
    license = Column(Integer)
    BQ2_GST = Column(Integer)
    BQ2_exGST = Column(Integer)
    TOTAL_BQ2 = Column(Integer)
    BQ2_rup = Column(Integer)
    plant_code = Column(String(255))
    purchase_group = Column(String(255))
    fund_centre = Column(String(255))
    BDP_clause = Column(String(255))
    PR_no = Column(String(255))
    RELEASE_STRAT = Column(String(255))

class Pricebid(Base):
    __tablename__ = "pricebid"
    case_id=Column(String(255))
    id = Column(Integer, primary_key=True)
    vendor = Column(String(255))
    add = Column(String(255))
    contact = Column(String(255))
    contact_person = Column(String(255))
    email = Column(String(255))
    license1 = Column(String(255))
    license2 = Column(String(255))
    license3 = Column(String(255))
    license1_no = Column(Integer)
    license2_no = Column(Integer)
    license3_no = Column(Integer)
    GST = Column(Integer)
    basis = Column(String(255))

class Priceschedule(Base):
    __tablename__ = "priceschedule"
    case_id=Column(String(255))
    id = Column(Integer, primary_key=True)
    license1 = Column(String(255))
    license2 = Column(String(255))
    license3 = Column(String(255))
    license4 = Column(String(255))
    license5 = Column(String(255))
    license6 = Column(String(255))
    license7 = Column(String(255))
    license8 = Column(String(255))
    license1_no = Column(Integer)
    license2_no = Column(Integer)
    license3_no = Column(Integer)
    license4_no = Column(Integer)
    license5_no = Column(Integer)
    license6_no = Column(Integer)
    license7_no = Column(Integer)
    license8_no = Column(Integer)
    GST = Column(Integer)
    basis = Column(String(255))

# Create tables
Base.metadata.create_all(engine)
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "new":
            return render_template("index.html", show_new_case=True, show_existing_case=False)
        elif action == "existing":
            return render_template("index.html", show_new_case=False, show_existing_case=True)

    return render_template("index.html")
@app.route("/download_template")
def download_template():
    return send_file("templates_excel/template_doc2.xlsx", as_attachment=True)
@app.route("/final_summary")
def final_summary():
    case_id = session.get("case_id")
    if not case_id:
        return "No case ID found. Please restart workflow.", 400

    # List all files from OUTPUT_FOLDER that start with case_id or contain it
    matching_files = [
        f for f in os.listdir(OUTPUT_FOLDER)
        if f.lower().startswith(case_id.lower()) or f"{case_id}_" in f
    ]

    return render_template("final_summary.html", case_id=case_id, files=matching_files)

# @app.route("/final_summary")
# def final_summary():
#     case_id = session.get("case_id")
#     output_files = session.get("output_files", [])

#     if not case_id:
#         return "No case ID found. Please restart workflow.", 400

#     filenames = [os.path.basename(f) for f in output_files]

#     return render_template("final_summary.html", case_id=case_id, files=filenames)

@app.route("/upload_case_data", methods=["POST"])
def upload_case_data():
    file = request.files.get("case_excel")
    if not file:
        return "No file uploaded", 400

    import pandas as pd
    df = pd.read_excel(file)
    try:
        extracted_case_id = df['case_id'].iloc[0]  # make sure your Excel has a column named 'case_id'
    except Exception as e:
        return "Error reading case_id from Excel. Ensure it's in the first row and column is named 'case_id'.", 400

    # Check if case_id already exists in UserDocs
    existing = db_session.query(UserDocs).filter_by(case_id=extracted_case_id).first()
    if existing:
        return render_template("index.html", error="Case ID already exists in UserDocs. Use a different file.", show_new_case=True)

    # If it's a new case_id, store in session temporarily and show document checklist
    session["case_id"] = extracted_case_id
    # Process DOC1 Excel upload
    if "case_excel" in request.files:
        doc1_file = request.files["case_excel"]
        if doc1_file.filename:
            output_files = process_doc2(
                doc1_file,
                db_session,
                RecordDoc2,
                UPLOAD_FOLDER,
                OUTPUT_FOLDER,
                "templates_word/template_doc2.docx",
                extracted_case_id
            )
            session["output_files"] = session.get("output_files", []) + output_files

    return render_template("index.html", show_doc_form=True, case_id=extracted_case_id, show_new_case=False, show_existing_case=False)
@app.route("/submit_documents", methods=["POST"])
def submit_documents():
    case_id = session.get("case_id")
    if not case_id:
        return "Session expired or case_id missing. Please upload Excel again.", 400

    doc1 = request.form.get("doc1")
    doc2 = request.form.get("doc2")
    doc3 = request.form.get("doc3")
    doc4 = request.form.get("doc4")

    new_entry = UserDocs(
        case_id=case_id,
        OEM_details=doc1 if doc1 else None,
        OEM_authorization=doc2 if doc2 else None,
        GAR=doc3 if doc3 else None,
        BQ=doc4 if doc4 else None
    )

    db_session.add(new_entry)
    db_session.commit()

    return redirect(url_for("upload_type_decider"))

# @app.route("/upload_case_data", methods=["POST"])
# def upload_case_data():
#     file = request.files.get("case_excel")
#     if not file:
#         return "No file uploaded", 400

#     case_id = datetime.now().strftime("CASE_%Y%m%d_%H%M%S")
#     session["case_id"] = case_id

#     # Call process_doc2 manually here
#     from logic.process_doc2 import process_doc2
#     output_files = process_doc2(file, db_session, RecordDoc2, UPLOAD_FOLDER, OUTPUT_FOLDER, "templates_word/template_doc2.docx", case_id)

#     return redirect(url_for("upload_excel"))

# @app.route("/", methods=["GET", "POST"])
# def index():
#     if request.method == "POST":
#         doc1 = request.form.get("doc1")
#         doc2 = request.form.get("doc2")
#         doc3 = request.form.get("doc3")
#         doc4 = request.form.get("doc4")

#         new_case_id = datetime.now().strftime("CASE_%Y%m%d_%H%M%S")
#         new_entry = UserDocs(
#             doc1=doc1 if doc1 else None,
#             doc2=doc2 if doc2 else None,
#             doc3=doc3 if doc3 else None,
#             doc4=doc4 if doc4 else None,
#             case_id=new_case_id
#         )
#         db_session.add(new_entry)
#         db_session.commit()

#         # Optional: store in session if you want to track last inserted record
#         session["user_id"] = new_entry.id
#         session["case_id"] = new_case_id

#         return redirect(url_for("upload_excel"))

#     return render_template("index.html", existing=None)

# @app.route("/upload_excel", methods=["GET", "POST"])
# def upload_excel():
#     if request.method == "POST":
#         output_files = []

#         file1 = request.files.get("file1")
#         file2 = request.files.get("file2")
#         file3 = request.files.get("file3")
#         file4 = request.files.get("file4")
#         case_id = session.get("case_id")

#         if file1:
#             output_files += process_doc1(file1, db_session, RecordDoc1,CaseDetailsDoc1, UPLOAD_FOLDER, OUTPUT_FOLDER, "templates_word/template_doc1.docx",case_id)

#         if file2:
#             output_files += process_doc2(file2, db_session, RecordDoc2, UPLOAD_FOLDER, OUTPUT_FOLDER, "templates_word/template_doc2.docx",case_id)

#         if file3:
#             output_files += process_doc3(file3, db_session, Pricebid, UPLOAD_FOLDER, OUTPUT_FOLDER, "templates_word/template_doc3.docx",case_id)

#         if file4:
#             output_files += process_doc4(file4, db_session, Priceschedule, UPLOAD_FOLDER, OUTPUT_FOLDER, "templates_word/template_doc4.docx",case_id)

#         filenames = [os.path.basename(f) for f in output_files]
#         return render_template("upload_excel.html", message="Documents Generated!", files=filenames,case_id=case_id)

#     return render_template("upload_excel.html", message=None, files=[])

@app.route("/download/<filename>")
def download(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

@app.route("/view_case_redirect", methods=["POST"])
def view_case_redirect():
    case_id = request.form.get("case_id")
    return redirect(url_for("view_case", case_id=case_id))
@app.template_filter('attr')
def attr(obj, attr_name):
    return getattr(obj, attr_name, "")
@app.route("/upload_type_decider", methods=["GET", "POST"])
def upload_type_decider():
    if request.method == "POST":
        type_file = request.files.get("type_file")
        if not type_file:
            return "No file uploaded", 400

        try:
            # Save the file first
            filepath = os.path.join(UPLOAD_FOLDER, type_file.filename)
            type_file.save(filepath)
            
            # Read the saved file
            df = pd.read_excel(filepath)
            
            if "type" not in df.columns:
                return "Excel must contain a 'type' column", 400

            file_type = str(df.iloc[0]["type"]).strip().lower()
            case_id = session.get("case_id")

            if file_type in ["capital", "amc"]:
                # Pass the file path to process_doc2
                output_files = process_doc1(
                    filepath,  # Pass path instead of file object
                    db_session,
                    RecordDoc1,
                    CaseDetailsDoc1,
                    UPLOAD_FOLDER,
                    OUTPUT_FOLDER,
                    "templates_word/template_doc1.docx",
                    case_id
                )
                session["output_files"] = session.get("output_files", []) + output_files
                
                if file_type == "capital":
                    return redirect(url_for("upload_doc3"))
                else:
                    return redirect(url_for("upload_doc4"))
            else:
                return f"Unknown type '{file_type}'. Expected 'capital' or 'amc'", 400

        except Exception as e:
            return f"Error processing file: {str(e)}", 400

    return render_template("upload_type_decider.html")
@app.route("/upload_doc3", methods=["GET", "POST"])
def upload_doc3():
    if request.method == "POST":
        file = request.files.get("doc3_file")
        case_id = session.get("case_id")
        if not file or not case_id:
            return "Missing file or case_id", 400
        process_doc3(file, db_session, Pricebid, UPLOAD_FOLDER, OUTPUT_FOLDER, "templates_word/template_doc3.docx", case_id)
        return redirect(url_for("final_summary"))
    return render_template("upload_doc3.html")


@app.route("/upload_doc4", methods=["GET", "POST"])
def upload_doc4():
    if request.method == "POST":
        file = request.files.get("doc4_file")
        case_id = session.get("case_id")
        if not file or not case_id:
            return "Missing file or case_id", 400
        process_doc4(file, db_session, Priceschedule, UPLOAD_FOLDER, OUTPUT_FOLDER, "templates_word/template_doc4.docx", case_id)
        return redirect(url_for("final_summary"))
    return render_template("upload_doc4.html")

@app.route("/view_case/<case_id>", methods=["GET", "POST"])
def view_case(case_id):
    if request.method == "POST":
    # 1. Update UserDocs (only one row)
        userdocs = db_session.query(UserDocs).filter_by(case_id=case_id).first()
        if userdocs:
            for col in userdocs.__table__.columns:
                if col.name not in ["id", "case_id"]:
                    form_key = f"userdocs_{col.name}"
                    if form_key in request.form:
                        setattr(userdocs, col.name, request.form.get(form_key))

    # 2. Update RecordDoc1 rows
        for row in db_session.query(RecordDoc1).filter_by(case_id=case_id).all():
            for col in row.__table__.columns:
                if col.name not in ["id", "case_id"]:
                    form_key = f"doc1_{row.id}_{col.name}"
                    if form_key in request.form:
                        setattr(row, col.name, request.form.get(form_key))

    # 3. Update RecordDoc2 rows
        for row in db_session.query(CaseDetailsDoc1).filter_by(case_id=case_id).all():
            for col in row.__table__.columns:
                if col.name not in ["id", "case_id"]:
                    form_key = f"doc1_{row.id}_{col.name}"
                    if form_key in request.form:
                        setattr(row, col.name, request.form.get(form_key))
        for row in db_session.query(RecordDoc2).filter_by(case_id=case_id).all():
            for col in row.__table__.columns:
                if col.name not in ["id", "case_id"]:
                    form_key = f"doc2_{row.id}_{col.name}"
                    if form_key in request.form:
                        setattr(row, col.name, request.form.get(form_key))
    # 4. Update Pricebid rows
        for row in db_session.query(Pricebid).filter_by(case_id=case_id).all():
            for col in row.__table__.columns:
                if col.name not in ["id", "case_id"]:
                    form_key = f"pricebid_{row.id}_{col.name}"
                    if form_key in request.form:
                        setattr(row, col.name, request.form.get(form_key))

    # 5. Update Priceschedule rows
        for row in db_session.query(Priceschedule).filter_by(case_id=case_id).all():
            for col in row.__table__.columns:
                if col.name not in ["id", "case_id"]:
                    form_key = f"pricesched_{row.id}_{col.name}"
                    if form_key in request.form:
                        setattr(row, col.name, request.form.get(form_key))

        db_session.commit()
        return redirect(url_for("view_case", case_id=case_id))


    # Load all rows (except UserDocs, which is one per case)
    userdocs = db_session.query(UserDocs).filter_by(case_id=case_id).first()
    doc1_list = db_session.query(RecordDoc1).filter_by(case_id=case_id).all()
    doc2_list = db_session.query(RecordDoc2).filter_by(case_id=case_id).all()
    desc_list = db_session.query(CaseDetailsDoc1).filter_by(case_id=case_id).all()

    pricebid_list = db_session.query(Pricebid).filter_by(case_id=case_id).all()
    pricesched_list = db_session.query(Priceschedule).filter_by(case_id=case_id).all()

    if not any([userdocs, doc1_list, doc2_list,desc_list, pricebid_list, pricesched_list]):
        return render_template("case_view.html", message="No records found for this Case ID.", case_id=case_id)

    return render_template("case_view.html",
                           case_id=case_id,
                           userdocs=userdocs,
                           doc1_list=doc1_list,
                           doc2_list=doc2_list,
                           desc_list=desc_list,
                           pricebid_list=pricebid_list,
                           pricesched_list=pricesched_list)


@app.teardown_appcontext
def shutdown_session(exception=None):
    db_session.close()

if __name__ == "__main__":
    app.run(debug=True)
