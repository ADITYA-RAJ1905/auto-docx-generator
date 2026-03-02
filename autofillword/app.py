from flask import Flask, render_template, request, send_file
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String, Date
from sqlalchemy.orm import sessionmaker, declarative_base
from docxtpl import DocxTemplate
import os
import datetime

app = Flask(__name__)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")


os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Database setup
engine = create_engine("sqlite:///data.db")
Session = sessionmaker(bind=engine)
session = Session()
Base = declarative_base()

class Record(Base):
    __tablename__ = "records"
    id = Column(Integer, primary_key=True)
    software_module = Column(String)
    address = Column(String)          # address
    desc = Column(String)
    desc1 = Column(String)
    desc2 = Column(String)
    L1 = Column(String)
    L2 = Column(String)
    L3 = Column(String)

Base.metadata.create_all(engine)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["excel_file"]
        if file:
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            df = pd.read_excel(filepath)
            session.query(Record).delete()

            for _, row in df.iterrows():
                record = Record(
                    software_module=row["Software_Module"],
                    address=row["address"],
                    desc=row["desc"],
                    desc1=row["desc1"],
                    desc2=row["desc2"],
                    L1=row["L1"],
                    L2=row["L2"],
                    L3=row["L3"],
                )
                session.add(record)
            session.commit()
            include_payment = 'include_payment' in request.form
            include_invoice = 'include_invoice' in request.form
            include_contractor = 'include_contractor' in request.form
            include_warranty = 'include_warranty' in request.form
            # Generate Word docs
            template = DocxTemplate(TEMPLATE_PATH)
            output_files = []

            for record in session.query(Record).all():
                context = {
                    "software_module": record.software_module,
                    "address":record.address,
                    "desc":record.desc,
                    "desc1":record.desc1,
                    "desc2":record.desc2,
                    "L1": record.L1,
                    "L2": record.L2,
                    "L3": record.L3,
                    "include_payment": include_payment,
                    "include_invoice": include_invoice,
                    "include_contractor": include_contractor,
                    "include_warranty": include_warranty
                }
                out_path = os.path.join(OUTPUT_FOLDER, f"{record.software_module}_output.docx")
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