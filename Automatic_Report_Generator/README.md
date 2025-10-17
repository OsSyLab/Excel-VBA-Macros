# 🧩 Advanced Search & Filter Macro  
**Excel VBA Macro #20 – Multi-Criteria Filtering Tool**

---

## 📄 Description
This macro allows users to filter Excel data dynamically based on multiple criteria such as **Department**, **Region**, and **Status**.  
It reads the user’s input from a small criteria panel and automatically generates a new sheet called **FilteredResults**, containing only the rows that match.

---

## ⚙️ How It Works
1. Open **Advanced_Search&Filter_Demo.xlsm**  
2. On the main sheet (`DataSheet`):
   - Enter filter criteria in the **Filter Criteria** panel (cells `H2:H4`).
   - Click **“Filter Data”** to run the filter.  
   - Optionally, click **“Clear Results”** to delete the filtered sheet and reset all filters.
3. The macro copies the headers, keeps column widths, and performs **case-insensitive partial matching**.

---

## 🧭 Example Screenshots
### 🔹 Before Filtering  
The user enters the filter criteria on the right panel, then clicks **Filter Data**.  
Only rows that meet all criteria will be extracted automatically.

### 🔹 After Filtering  
A new sheet called **FilteredResults** is created.  
It displays only matching rows while maintaining consistent formatting.  
Clicking **Clear Results** completely deletes the *FilteredResults* sheet and restores *DataSheet* to its clean state.

---

## 🧠 Technical Details
- **VBA Module:** `Advanced_Search&Filter.bas`  
- **Main Procedure:** `Run_AdvancedFilter()`  
- **Reset Procedure:** `Clear_FilterResults()`  
- **Filter Logic:** Uses `AutoFilter` with wildcards (`=*criteria*`) for flexible search.  
- **Output Sheet:** Automatically generated (`FilteredResults`).

---

## 🧩 File Structure
Advanced_Search&Filter.bas
Advanced_Search&Filter_Demo.xlsm
Screenshot_1.png

---

## 🪄 Customization Tips
- To switch from **partial match** to **exact match**, replace  
  ```vb
  Criteria1:="=*" & critDept & "*"
with


Criteria1:=critDept
You can change filter fields or output sheet name directly inside the VBA code.

License

**MIT License**  
You are free to use, modify, and distribute this code with attribution.  

© 2025 **Data Solutions Lab. by Osman Uluhan** – All rights reserved.
