# Auto DOCX Generator

An internal automation tool built using Flask and MySQL to generate standardized Word (DOCX) reports from structured Excel data. The system was designed to streamline repetitive documentation workflows by dynamically populating predefined templates and maintaining case records for future retrieval.

---

## 🚀 Overview

Organizations often rely on standardized Word templates where only specific fields change per case. Manually editing these templates is time-consuming and error-prone.

This system automates that process by:

- Accepting structured Excel input files
- Extracting relevant data using pandas
- Populating predefined DOCX templates dynamically
- Generating 4–5 customized reports per case
- Storing case data in MySQL for future retrieval

---

## 🛠 Tech Stack

- Python
- Flask
- MySQL
- pandas
- python-docx
- SQL

---

## ⚙️ Features

- 📥 Upload structured Excel files (.xlsx)
- 📝 Automatically populate Word templates with dynamic field values
- 📄 Generate multiple DOCX reports based on user selection
- 🗂 Maintain case records using MySQL (multiple relational tables)
- 🔎 Retrieve previously generated documents using Case ID
- 🖥 Local deployment for internal organizational use

---

## 🏗 System Architecture

1. User uploads structured Excel file
2. Data is processed using pandas
3. Extracted values are inserted into predefined DOCX templates
4. Multiple reports are generated dynamically
5. Case data is stored in MySQL for tracking and retrieval

---

---

## 📈 Impact

- Reduced repetitive manual documentation effort
- Improved consistency and accuracy in standardized reports
- Enabled structured case tracking and document retrieval
- Used internally for streamlined workflow automation

---

## 🔮 Future Improvements

- Add authentication and role-based access control
- Deploy on cloud infrastructure
- Add logging and audit tracking
- Implement PDF export support
