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

class Record(Base):
    __tablename__ = "records"
    id = Column(Integer, primary_key=True)
    tag = Column(String(255))
    subject=Column(String(255))
    file_no=Column(String(255))
    material=Column(String(255))
    vendor=Column(String(255))
    tender_type=Column(String(255))
    user_disha_file=Column(String(255))
    BQ1_date = Column(String(255)) 
    Proposal_curr = Column(String(255))
    BQ1_price = Column(Integer)
    BQ2_price=Column(Integer)
    LPR_PO = Column(Integer)
    BQ_per = Column(Integer)
    FY = Column(String(255))
    CURR_EXC_RATE=Column(Integer)
    LPR_UNIT_PRICE = Column(Integer)
    PRICE_DIFF=Column(Integer)
    license=Column(Integer)
    BQ2_GST=Column(Integer)
    BQ2_exGST=Column(Integer)
    TOTAL_BQ2=Column(Integer)
    BQ2_rup=Column(Integer)
    plant_code=Column(String(255))
    purchase_group=Column(String(255))
    fund_centre=Column(String(255))
    BDP_clause=Column(String(255))
    PR_no=Column(String(255))
    RELEASE_STRAT=Column(String(255))
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
                # Safe default function
                def safe_round(value):
                    if value is None or (isinstance(value, float) and math.isnan(value)):
                        return 0
                    return round(value)

                price_diff = row["BQ2_price"] - row["BQ1_price"] if not pd.isna(row["BQ2_price"]) and not pd.isna(row["BQ1_price"]) else 0

                bq1_price = row["BQ1_price"] if not pd.isna(row["BQ1_price"]) else 0
                bq2_price = row["BQ2_price"] if not pd.isna(row["BQ2_price"]) else 0
                curr_exc_rate = row["CURR_EXC_RATE"] if not pd.isna(row["CURR_EXC_RATE"]) else 0
                lic=row["license"] if not pd.isna(row["license"]) else 0
                bq_per = (bq2_price / bq1_price) * 100 if bq1_price else 0
                bq2_gst = bq2_price + (18/100 * bq2_price) if bq2_price else 0
                bq2_exgst=lic*bq2_price
                total_bq2=lic*bq2_gst
                bq2_rup = total_bq2 * curr_exc_rate if curr_exc_rate else 0
                record = Record(
                    tag=row["tag"],
                    subject=row["subject"],
                    file_no=row["file_no"],
                    material=row["material"],
                    vendor=row["vendor"],
                    tender_type=row["tender_type"],
                    user_disha_file=row["user_disha_file"],
                    BQ1_date=row["BQ1_date"],
                    Proposal_curr=row["Proposal_curr"],
                    BQ1_price=safe_round(bq1_price),
                    BQ2_price=safe_round(bq2_price),
                    LPR_PO=safe_round(row["LPR_PO"]) if not pd.isna(row["LPR_PO"]) else 0,
                    BQ_per=safe_round(bq_per),
                    FY=row["FY"],
                    CURR_EXC_RATE=safe_round(curr_exc_rate),
                    LPR_UNIT_PRICE=safe_round(row["LPR_UNIT_PRICE"]) if not pd.isna(row["LPR_UNIT_PRICE"]) else 0,
                    PRICE_DIFF=safe_round(price_diff),
                    license=safe_round(row["license"]) if not pd.isna(row["license"]) else 0,
                    BQ2_GST=safe_round(bq2_gst),
                    BQ2_exGST=safe_round(bq2_exgst),
                    TOTAL_BQ2=safe_round(total_bq2),
                    BQ2_rup=safe_round(bq2_rup),
                    plant_code=row["plant_code"],
                    purchase_group=row["purchase_group"],
                    fund_centre=row["fund_centre"],
                    BDP_clause=row["BDP_clause"],
                    PR_no=row["PR_no"],
                    RELEASE_STRAT=row["RELEASE_STRAT"],
                )
                session.add(record)

            session.commit()

            # Generate Word documents
            template = DocxTemplate(TEMPLATE_PATH)
            output_files = []

            for record in session.query(Record).all():
                context = {
                    "tag": record.tag,
                    "subject": record.subject,
                    "file_no": record.file_no,
                    "material": record.material,
                    "vendor": record.vendor,
                    "tender_type": record.tender_type,
                    "user_disha_file": record.user_disha_file,
                    "BQ1_date": record.BQ1_date,
                    "Proposal_curr": record.Proposal_curr,
                    "BQ1_price": record.BQ1_price,
                    "BQ2_price": record.BQ2_price,
                    "LPR_PO": record.LPR_PO,
                    "BQ_per": record.BQ_per,
                    "FY": record.FY,
                    "CURR_EXC_RATE": record.CURR_EXC_RATE,
                    "LPR_UNIT_PRICE": record.LPR_UNIT_PRICE,
                    "PRICE_DIFF": record.PRICE_DIFF,
                    "license": record.license,
                    "BQ2_GST": record.BQ2_GST,
                    "BQ2_exGST": record.BQ2_exGST,
                    "TOTAL_BQ2": record.TOTAL_BQ2,
                    "BQ2_rup": record.BQ2_rup,
                    "plant_code": record.plant_code,
                    "purchase_group": record.purchase_group,
                    "fund_centre": record.fund_centre,
                    "BDP_clause": record.BDP_clause,
                    "PR_no": record.PR_no,
                    "RELEASE_STRAT": record.RELEASE_STRAT
                }

                out_path = os.path.join(OUTPUT_FOLDER, f"Proposal_for_procurement_of_{record.material}_{record.vendor}_{record.tender_type}.docx")
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