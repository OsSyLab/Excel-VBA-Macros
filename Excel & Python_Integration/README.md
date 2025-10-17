# 🐍 Excel & Python Integration

**Excel VBA Macro #19 – Run Python Scripts Directly from Excel**

This macro enables **seamless integration between Excel and Python** — allowing users to send data from Excel to Python, process it there, and return the results automatically.  
It’s the perfect bridge for combining Excel’s interface with Python’s analytical and automation power.

---

## 📁 Files Included

| File | Description |
|------|--------------|
| `1.Excel_Python_Integration.bas` | VBA module to execute Python scripts from Excel |
| `2.Excel_Python_Integration_Demo.xlsm` | Demo workbook showing the Python connection workflow |
| `3.Excel_Python_Integration_Readme.md` | User guide and integration steps |
| `4.Excel_Python_Integration_Screenshot.png` | Screenshot showing the Excel–Python data flow |

---

## ⚙️ How It Works

1. The macro exports a selected Excel range to a temporary CSV file.  
2. Calls a **Python script** (via command line or direct path).  
3. Python processes the data (e.g., analysis, predictions, transformations).  
4. The macro imports the resulting data back into Excel automatically.  

---

## 🧠 Example Use Case

- Automate data cleaning or analysis with Python and view results in Excel  
- Run machine learning predictions directly from Excel sheets  
- Integrate Excel dashboards with Python analytics pipelines  

---

## ⚠️ Setup Notes

- Python must be **installed and added to PATH**  
- Required Python libraries (e.g., pandas, openpyxl) should be installed  
- You can customize script and folder paths inside the VBA module  

---

## 🧾 Requirements

- Microsoft Excel **2013 or newer**  
- VBA enabled  
- Python 3.x installed  

---

## 🖼️ Preview
*(See included screenshot for workflow demonstration.)*

---

📂 **Path:**  
`Excel-VBA-Macros/Excel_Python_Integration/`

## License
You are free to use, modify, and distribute this code with attribution.
© 2025 Data Solutions Lab. by Osman Uluhan – All rights reserved.
