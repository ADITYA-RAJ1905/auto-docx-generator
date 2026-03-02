import os
import pandas as pd
from docxtpl import DocxTemplate
from flask import session as flask_session

def process_doc4(file, db_session, Priceschedule, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH, case_id):
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    df = pd.read_excel(filepath)

    def safe_round(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return 0
        return round(value)

    for _, row in df.iterrows():
        priceschedule = Priceschedule(
            license1=row["license1"],
            license2=row["license2"],
            license3=row["license3"],
            license4=row["license4"],
            license5=row["license5"],
            license6=row["license6"],
            license7=row["license7"],
            license8=row["license8"],
            license1_no=safe_round(row["license1_no"]),
            license2_no=safe_round(row["license2_no"]),
            license3_no=safe_round(row["license3_no"]),
            license4_no=safe_round(row["license4_no"]),
            license5_no=safe_round(row["license5_no"]),
            license6_no=safe_round(row["license6_no"]),
            license7_no=safe_round(row["license7_no"]),
            license8_no=safe_round(row["license8_no"]),
            GST=safe_round(row["GST"]),
            basis=row["basis"],
            case_id=case_id
        )
        db_session.add(priceschedule)
    db_session.commit()

    template = DocxTemplate(TEMPLATE_PATH)
    output_files = []

    for ps in db_session.query(Priceschedule).filter_by(case_id=case_id).all():
        context = {
            "license1": ps.license1,
            "license2": ps.license2,
            "license3": ps.license3,
            "license4": ps.license4,
            "license5": ps.license5,
            "license6": ps.license6,
            "license7": ps.license7,
            "license8": ps.license8,
            "license1_no": ps.license1_no,
            "license2_no": ps.license2_no,
            "license3_no": ps.license3_no,
            "license4_no": ps.license4_no,
            "license5_no": ps.license5_no,
            "license6_no": ps.license6_no,
            "license7_no": ps.license7_no,
            "license8_no": ps.license8_no,
            "GST": ps.GST,
            "basis": ps.basis,
            "case_id": case_id
        }
        out_path = os.path.join(OUTPUT_FOLDER, f"{case_id}_SOW_AMC_{case_id}.docx")
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

# def process_doc4(file, session, Priceschedule, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH,case_id):
#     filepath = os.path.join(UPLOAD_FOLDER, file.filename)
#     file.save(filepath)
#     df = pd.read_excel(filepath)

#     def safe_round(value):
#         if value is None or (isinstance(value, float) and pd.isna(value)):
#             return 0
#         return round(value)

#     for _, row in df.iterrows():
#         priceschedule = Priceschedule(
#             license1=row["license1"],
#             license2=row["license2"],
#             license3=row["license3"],
#             license4=row["license4"],
#             license5=row["license5"],
#             license6=row["license6"],
#             license7=row["license7"],
#             license8=row["license8"],
#             license1_no=safe_round(row["license1_no"]),
#             license2_no=safe_round(row["license2_no"]),
#             license3_no=safe_round(row["license3_no"]),
#             license4_no=safe_round(row["license4_no"]),
#             license5_no=safe_round(row["license5_no"]),
#             license6_no=safe_round(row["license6_no"]),
#             license7_no=safe_round(row["license7_no"]),
#             license8_no=safe_round(row["license8_no"]),
#             GST=safe_round(row["GST"]),
#             basis=row["basis"],
#             case_id=case_id
#         )
#         session.add(priceschedule)
#     session.commit()

#     template = DocxTemplate(TEMPLATE_PATH)
#     output_files = []

#     for priceschedule in session.query(Priceschedule).filter_by(case_id=case_id).all():
#         context = {
#             "license1": priceschedule.license1,
#             "license2": priceschedule.license2,
#             "license3": priceschedule.license3,
#             "license4": priceschedule.license4,
#             "license5": priceschedule.license5,
#             "license6": priceschedule.license6,
#             "license7": priceschedule.license7,
#             "license8": priceschedule.license8,
#             "license1_no": priceschedule.license1_no,
#             "license2_no": priceschedule.license2_no,
#             "license3_no": priceschedule.license3_no,
#             "license4_no": priceschedule.license4_no,
#             "license5_no": priceschedule.license5_no,
#             "license6_no": priceschedule.license6_no,
#             "license7_no": priceschedule.license7_no,
#             "license8_no": priceschedule.license8_no,
#             "GST": priceschedule.GST,
#             "basis": priceschedule.basis,
#             "case_id":case_id
#         }
#         out_path = os.path.join(OUTPUT_FOLDER, f"SOW_AMC.docx")
#         template.render(context)
#         template.save(out_path)
#         output_files.append(out_path)

#     return output_files
