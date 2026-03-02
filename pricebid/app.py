from flask import Flask, render_template, request, send_file
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String, Date
from sqlalchemy.orm import sessionmaker, declarative_base
from docxtpl import DocxTemplate
import os
import datetime
import math
app = Flask(__name__)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")


os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Database setup
engine = create_engine("mysql+pymysql://root:Vidyad%401905@localhost/proposal")
Session = sessionmaker(bind=engine)
session = Session()
Base = declarative_base()

class Pricebid(Base):
    __tablename__ = "pricebid"
    id = Column(Integer, primary_key=True)
    vendor = Column(String(255))
    add=Column(String(255))
    contact=Column(String(255))
    contact_person=Column(String(255))
    email=Column(String(255))
    license1=Column(String(255))
    license2=Column(String(255))
    license3 = Column(String(255)) 
    license1_no = Column(Integer)
    license2_no=Column(Integer)
    license3_no=Column(Integer)
    GST=Column(Integer)
    basis=Column(String(255))
Base.metadata.create_all(engine)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["excel_file"]
        if file:
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            df = pd.read_excel(filepath)
            session.query(Pricebid).delete()

            for _, row in df.iterrows():
                # Safe default function
                def safe_round(value):
                    if value is None or (isinstance(value, float) and math.isnan(value)):
                        return 0
                    return round(value)
                pricebid = Pricebid(
                    vendor = row["vendor"],
                    add = row["add"],
                    contact = row["contact"],
                    contact_person = row["contact_person"],
                    email = row["email"],
                    license1 = row["license1"],
                    license2 = row["license2"],
                    license3 = row["license3"],
                    license1_no = row["license1_no"],
                    license2_no = row["license2_no"],
                    license3_no = row["license3_no"],
                    GST = row["GST"],
                    basis = row["basis"]

                )
                session.add(pricebid)

            session.commit()

            # Generate Word documents
            template = DocxTemplate(TEMPLATE_PATH)
            output_files = []

            for pricebid in session.query(Pricebid).all():
                context = {
                    "vendor": pricebid.vendor,
                    "add": pricebid.add,
                    "contact": pricebid.contact,
                    "contact_person": pricebid.contact_person,
                    "email": pricebid.email,
                    "license1": pricebid.license1,
                    "license2": pricebid.license2,
                    "license3": pricebid.license3,
                    "license1_no": pricebid.license1_no,
                    "license2_no": pricebid.license2_no,
                    "license3_no": pricebid.license3_no,
                    "GST": pricebid.GST,
                    "basis": pricebid.basis

                }

                out_path = os.path.join(OUTPUT_FOLDER, f"Price Bid.docx")
                template.render(context)
                template.save(out_path)
                output_files.append(out_path)

            return render_template("index.html", message="Success! Documents generated.", files=output_files)

    return render_template("index.html", message=None, files=[])


@app.route("/download/<filename>")
def download(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True) 