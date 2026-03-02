import os
import math
import pandas as pd
from sqlalchemy.orm import sessionmaker
from docxtpl import DocxTemplate
from flask import session as flask_session  # import at top
import re
def sanitize_filename(name):
    """Sanitize and fallback to 'record' if name is bad."""
    name = str(name).strip() if name else "record"
    name = re.sub(r"[\\/*?\"<>|:\n\r\t]", "_", name)
    return name[:50] or "record"  # truncate to 50 chars max

def process_doc2(file_or_path, db_session, Record, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH, case_id):
    if hasattr(file_or_path, 'filename'):  # It's a file object
        filepath = os.path.join(UPLOAD_FOLDER, file_or_path.filename)
        file_or_path.save(filepath)
    else:  # It's a path string
        filepath = file_or_path

    try:
        df = pd.read_excel(file_or_path)

        def safe_round(value):
            if value is None or (isinstance(value, float) and pd.isna(value)):
                return 0
            return round(value)

        for _, row in df.iterrows():
            price_diff = (row["BQ2_price"] - row["BQ1_price"]) if not pd.isna(row["BQ2_price"]) and not pd.isna(row["BQ1_price"]) else 0
            bq1_price = row["BQ1_price"] if not pd.isna(row["BQ1_price"]) else 0
            bq2_price = row["BQ2_price"] if not pd.isna(row["BQ2_price"]) else 0
            curr_exc_rate = row["CURR_EXC_RATE"] if not pd.isna(row["CURR_EXC_RATE"]) else 0
            lic = row["license"] if not pd.isna(row["license"]) else 0
            bq_per = (bq2_price / bq1_price) * 100 if bq1_price else 0
            bq2_gst = bq2_price + (0.18 * bq2_price) if bq2_price else 0
            bq2_exgst = lic * bq2_price
            total_bq2 = lic * bq2_gst
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
                license=safe_round(lic),
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
                case_id=case_id
            )
            db_session.add(record)
        db_session.commit()

        template = DocxTemplate(TEMPLATE_PATH)
        output_files = []

        for record in db_session.query(Record).filter_by(case_id=case_id).all():
            context = {col.name: getattr(record, col.name) for col in Record.__table__.columns if col.name != 'id'}
            safe_material = sanitize_filename(record.material)
            safe_vendor = sanitize_filename(record.vendor)
            safe_tender = sanitize_filename(record.tender_type)
            out_path = os.path.join(OUTPUT_FOLDER, f"{case_id}_Proposal_{safe_material}_{safe_vendor}_{safe_tender}.docx")

            template.render(context)
            template.save(out_path)
            output_files.append(out_path)

        # ✅ Store output files in Flask session
        existing_files = flask_session.get("output_files", [])
        existing_files+=output_files
        #flask_session["output_files"] = existing_files

        return output_files

    except Exception as e:
        raise Exception(f"Error processing doc2: {str(e)}")

# import os
# import math
# import pandas as pd
# from sqlalchemy.orm import sessionmaker
# from docxtpl import DocxTemplate
# from flask import session as flask_session  # import at top

# def process_doc2(file, db_session, Record, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH,case_id):
#     filepath = os.path.join(UPLOAD_FOLDER, file.filename)
#     file.save(filepath)
#     df = pd.read_excel(filepath)

#     def safe_round(value):
#         if value is None or (isinstance(value, float) and pd.isna(value)):
#             return 0
#         return round(value)

#     for _, row in df.iterrows():
#         price_diff = (row["BQ2_price"] - row["BQ1_price"]) if not pd.isna(row["BQ2_price"]) and not pd.isna(row["BQ1_price"]) else 0
#         bq1_price = row["BQ1_price"] if not pd.isna(row["BQ1_price"]) else 0
#         bq2_price = row["BQ2_price"] if not pd.isna(row["BQ2_price"]) else 0
#         curr_exc_rate = row["CURR_EXC_RATE"] if not pd.isna(row["CURR_EXC_RATE"]) else 0
#         lic = row["license"] if not pd.isna(row["license"]) else 0
#         bq_per = (bq2_price / bq1_price) * 100 if bq1_price else 0
#         bq2_gst = bq2_price + (18/100 * bq2_price) if bq2_price else 0
#         bq2_exgst = lic * bq2_price
#         total_bq2 = lic * bq2_gst
#         bq2_rup = total_bq2 * curr_exc_rate if curr_exc_rate else 0

#         record = Record(
#             tag=row["tag"],
#             subject=row["subject"],
#             file_no=row["file_no"],
#             material=row["material"],
#             vendor=row["vendor"],
#             tender_type=row["tender_type"],
#             user_disha_file=row["user_disha_file"],
#             BQ1_date=row["BQ1_date"],
#             Proposal_curr=row["Proposal_curr"],
#             BQ1_price=safe_round(bq1_price),
#             BQ2_price=safe_round(bq2_price),
#             LPR_PO=safe_round(row["LPR_PO"]) if not pd.isna(row["LPR_PO"]) else 0,
#             BQ_per=safe_round(bq_per),
#             FY=row["FY"],
#             CURR_EXC_RATE=safe_round(curr_exc_rate),
#             LPR_UNIT_PRICE=safe_round(row["LPR_UNIT_PRICE"]) if not pd.isna(row["LPR_UNIT_PRICE"]) else 0,
#             PRICE_DIFF=safe_round(price_diff),
#             license=safe_round(lic),
#             BQ2_GST=safe_round(bq2_gst),
#             BQ2_exGST=safe_round(bq2_exgst),
#             TOTAL_BQ2=safe_round(total_bq2),
#             BQ2_rup=safe_round(bq2_rup),
#             plant_code=row["plant_code"],
#             purchase_group=row["purchase_group"],
#             fund_centre=row["fund_centre"],
#             BDP_clause=row["BDP_clause"],
#             PR_no=row["PR_no"],
#             RELEASE_STRAT=row["RELEASE_STRAT"],
#             case_id=case_id
#         )
#         db_session.add(record)
#     db_session.commit()

#     template = DocxTemplate(TEMPLATE_PATH)
#     output_files = []

#     for record in db_session.query(Record).filter_by(case_id=case_id).all():
#         context = {col.name: getattr(record, col.name) for col in Record.__table__.columns if col.name != 'id'}
#         out_path = os.path.join(OUTPUT_FOLDER, f"Proposal_{record.material}_{record.vendor}_{record.tender_type}.docx")
#         template.render(context)
#         template.save(out_path)
#         output_files.append(out_path)
#                 # After calling process_docX
#         existing_files = db_session.get("output_files", [])
#         existing_files += output_files
#         db_session["output_files"] = existing_files

#     return output_files
