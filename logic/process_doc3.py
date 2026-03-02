import os
import pandas as pd
from docxtpl import DocxTemplate
from flask import session as flask_session

def process_doc3(file, db_session, Pricebid, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH, case_id):
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    df = pd.read_excel(filepath)

    def safe_round(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return 0
        return round(value)

    for _, row in df.iterrows():
        pricebid = Pricebid(
            vendor=row["vendor"],
            add=row["add"],
            contact=row["contact"],
            contact_person=row["contact_person"],
            email=row["email"],
            license1=row["license1"],
            license2=row["license2"],
            license3=row["license3"],
            license1_no=safe_round(row["license1_no"]),
            license2_no=safe_round(row["license2_no"]),
            license3_no=safe_round(row["license3_no"]),
            GST=safe_round(row["GST"]),
            basis=row["basis"],
            case_id=case_id
        )
        db_session.add(pricebid)
    db_session.commit()

    template = DocxTemplate(TEMPLATE_PATH)
    output_files = []

    for pricebid in db_session.query(Pricebid).filter_by(case_id=case_id).all():
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
            "basis": pricebid.basis,
            "case_id": case_id
        }
        out_path = os.path.join(OUTPUT_FOLDER, f"{case_id}_SOW_CAPITAL_{pricebid.vendor}.docx")
        template.render(context)
        template.save(out_path)
        output_files.append(out_path)

    # ✅ Track in session
    existing_files = flask_session.get("output_files", [])
    #flask_session["output_files"] = existing_files + output_files

    return output_files

# import os
# import math
# import pandas as pd
# from docxtpl import DocxTemplate
# from flask import session as flask_session  # import at top

# def process_doc3(file, session, Pricebid, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH,case_id):
#     filepath = os.path.join(UPLOAD_FOLDER, file.filename)
#     file.save(filepath)
#     df = pd.read_excel(filepath)

#     def safe_round(value):
#         if value is None or (isinstance(value, float) and pd.isna(value)):
#             return 0
#         return round(value)

#     for _, row in df.iterrows():
#         pricebid = Pricebid(
#             vendor=row["vendor"],
#             add=row["add"],
#             contact=row["contact"],
#             contact_person=row["contact_person"],
#             email=row["email"],
#             license1=row["license1"],
#             license2=row["license2"],
#             license3=row["license3"],
#             license1_no=safe_round(row["license1_no"]),
#             license2_no=safe_round(row["license2_no"]),
#             license3_no=safe_round(row["license3_no"]),
#             GST=safe_round(row["GST"]),
#             basis=row["basis"],
#             case_id=case_id,
#         )
#         session.add(pricebid)
#     session.commit()

#     template = DocxTemplate(TEMPLATE_PATH)
#     output_files = []

#     for pricebid in session.query(Pricebid).filter_by(case_id=case_id).all():
#         context = {
#             "vendor": pricebid.vendor,
#             "add": pricebid.add,
#             "contact": pricebid.contact,
#             "contact_person": pricebid.contact_person,
#             "email": pricebid.email,
#             "license1": pricebid.license1,
#             "license2": pricebid.license2,
#             "license3": pricebid.license3,
#             "license1_no": pricebid.license1_no,
#             "license2_no": pricebid.license2_no,
#             "license3_no": pricebid.license3_no,
#             "GST": pricebid.GST,
#             "basis": pricebid.basis,
#             "case_id":case_id
#         }
#         out_path = os.path.join(OUTPUT_FOLDER, f"SOW_CAPITAL.docx")
#         template.render(context)
#         template.save(out_path)
#         output_files.append(out_path)

#     return output_files
