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

class Priceschedule(Base):
    __tablename__ = "priceschedule"
    id = Column(Integer, primary_key=True)
    license1=Column(String(255))
    license2=Column(String(255))
    license3 = Column(String(255)) 
    license4=Column(String(255))
    license5=Column(String(255))
    license6 = Column(String(255))
    license7 = Column(String(255))
    license8 = Column(String(255))

    license1_no = Column(Integer)
    license2_no=Column(Integer)
    license3_no=Column(Integer)
    license4_no = Column(Integer)
    license5_no=Column(Integer)
    license6_no=Column(Integer)
    license7_no=Column(Integer)
    license8_no=Column(Integer)

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
            session.query(Priceschedule).delete()

            for _, row in df.iterrows():
                # Safe default function
                def safe_round(value):
                    if value is None or (isinstance(value, float) and math.isnan(value)):
                        return 0
                    return round(value)
                priceschedule = Priceschedule(
                    license1 = row["license1"],
                    license2 = row["license2"],
                    license3 = row["license3"],
                    license4 = row["license4"],
                    license5 = row["license5"],
                    license6 = row["license6"],
                    license7 = row["license7"],
                    license8 = row["license8"],
                    license1_no = row["license1_no"],
                    license2_no = row["license2_no"],
                    license3_no = row["license3_no"],
                    license4_no = row["license4_no"],
                    license5_no = row["license5_no"],
                    license6_no = row["license6_no"],
                    license7_no = row["license7_no"],
                    license8_no = row["license8_no"],                    
                    GST = row["GST"],
                    basis = row["basis"]

                )
                session.add(priceschedule)

            session.commit()

            # Generate Word documents
            template = DocxTemplate(TEMPLATE_PATH)
            output_files = []

            for priceschedule in session.query(Priceschedule).all():
                context = {
                    "license1": priceschedule.license1,
                    "license2": priceschedule.license2,
                    "license3": priceschedule.license3,
                    "license4": priceschedule.license4,
                    "license5": priceschedule.license5,
                    "license6": priceschedule.license6,
                    "license7": priceschedule.license7,
                    "license8": priceschedule.license8,
                    "license1_no": priceschedule.license1_no,
                    "license2_no": priceschedule.license2_no,
                    "license3_no": priceschedule.license3_no,
                    "license4_no": priceschedule.license4_no,
                    "license5_no": priceschedule.license5_no,
                    "license6_no": priceschedule.license6_no,
                    "license7_no": priceschedule.license7_no,
                    "license8_no": priceschedule.license8_no,
                    "GST": priceschedule.GST,
                    "basis": priceschedule.basis

                }

                out_path = os.path.join(OUTPUT_FOLDER, f"Price Schedule.docx")
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