import os
import pandas as pd
from sqlalchemy.orm import sessionmaker
from docxtpl import DocxTemplate
from flask import session as flask_session  # Flask session
import re
def sanitize_filename(name):
    """Sanitize and fallback to 'record' if name is bad."""
    name = str(name).strip() if name else "record"
    name = re.sub(r"[\\/*?\"<>|:\n\r\t]", "_", name)
    return name[:50] or "record"  # truncate to 50 chars max

def process_doc1(file_or_path, db_session, Record, CaseDetailsDoc1, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH, case_id):
    if hasattr(file_or_path, 'filename'):  # It's a file object
        filepath = os.path.join(UPLOAD_FOLDER, file_or_path.filename)
        file_or_path.save(filepath)
    else:  # It's a path string
        filepath = file_or_path

    try:
        df = pd.read_excel(filepath)

        records = []
        details_list = []

        for _, row in df.iterrows():
            record = Record(
                software_module=row["Software_Module"],
                address=row["address"],
                l_valid=row["l_valid"],
                installation=row["installation"],
                certificate_days=row["certificate_days"],
                case_id=case_id
            )
            db_session.add(record)
            records.append(record)

            details = CaseDetailsDoc1(
                desc1=row["desc1"],
                desc2=row["desc2"],
                desc3=row["desc3"],
                L1=row["L1"],
                L2=row["L2"],
                L3=row["L3"],
                case_id=case_id
            )
            db_session.add(details)
            details_list.append(details)

        db_session.commit()  # commit to generate IDs if needed

        # Reload the data from DB to get all inserted rows (optional but safer)
        records_db = db_session.query(Record).filter_by(case_id=case_id).order_by(Record.id).all()
        details_db = db_session.query(CaseDetailsDoc1).filter_by(case_id=case_id).order_by(CaseDetailsDoc1.id).all()

        output_files = []

        # Assuming records_db and details_db are aligned 1:1 by order
        for record, details in zip(records_db, details_db):
            template = DocxTemplate(TEMPLATE_PATH)
            context = {
                "case_id": case_id,
                "software_module": record.software_module,
                "address": record.address,
                "l_valid":record.l_valid,
                "installation":record.installation,
                "certificate_days":record.certificate_days,
                "desc1": details.desc1 or "",
                "desc2": details.desc2 or "",
                "desc3": details.desc3 or "",
                "L1": details.L1 or "",
                "L2": details.L2 or "",
                "L3": details.L3 or "",
            }
            
            safe_module = sanitize_filename(record.software_module)
            out_path = os.path.join(OUTPUT_FOLDER, f"{case_id}_{safe_module}_{record.id}_output.docx")
            template.render(context)
            template.save(out_path)
            output_files.append(out_path)

        # ✅ Store generated files in Flask session
        existing_files = flask_session.get("output_files", [])
        existing_files+=output_files
        #flask_session["output_files"] = existing_files

        return output_files
    except Exception as e:
        raise Exception(f"Error processing doc2: {str(e)}")
# import os
# import pandas as pd
# from sqlalchemy.orm import sessionmaker
# from docxtpl import DocxTemplate

# def process_doc1(file, session, Record, CaseDetailsDoc1, UPLOAD_FOLDER, OUTPUT_FOLDER, TEMPLATE_PATH, options, case_id):
#     filepath = os.path.join(UPLOAD_FOLDER, file.filename)
#     file.save(filepath)
#     df = pd.read_excel(filepath)

#     records = []
#     details_list = []

#     for _, row in df.iterrows():
#         record = Record(
#             software_module=row["Software_Module"],
#             address=row["address"],
#             case_id=case_id
#         )
#         session.add(record)
#         records.append(record)

#         details = CaseDetailsDoc1(
#             desc1=row["desc1"],
#             desc2=row["desc2"],
#             desc3=row["desc3"],
#             L1=row["L1"],
#             L2=row["L2"],
#             L3=row["L3"],
#             case_id=case_id
#         )
#         session.add(details)
#         details_list.append(details)

#     session.commit()  # commit to generate IDs if needed

#     # Reload the data from DB to get all inserted rows (optional but safer)
#     records_db = session.query(Record).filter_by(case_id=case_id).order_by(Record.id).all()
#     details_db = session.query(CaseDetailsDoc1).filter_by(case_id=case_id).order_by(CaseDetailsDoc1.id).all()

#     output_files = []

#     # Assuming records_db and details_db are aligned 1:1 by order
#     for record, details in zip(records_db, details_db):
#         template = DocxTemplate(TEMPLATE_PATH)
#         context = {
#             "case_id": case_id,
#             "software_module": record.software_module,
#             "address": record.address,
#             "desc1": details.desc1 or "",
#             "desc2": details.desc2 or "",
#             "desc3": details.desc3 or "",
#             "L1": details.L1 or "",
#             "L2": details.L2 or "",
#             "L3": details.L3 or "",
#             **options
#         }
#         out_path = os.path.join(OUTPUT_FOLDER, f"{record.software_module}_output.docx")
#         template.render(context)
#         template.save(out_path)
#         output_files.append(out_path)

#     return output_files
