# KN-Logistics-Automation-Project
VBA automation solution created during a Data Operations internship at Kuehne + Nagel to efficiently generate filtered logistics reports from operational data
# KN-Logistics-Dashboard-VBA

This project showcases an operational **Excel Dashboard** and its underlying **VBA automation** developed during my Data Operations Internship at **Kuehne + Nagel**.

The solution provides a consolidated view of critical logistics KPIs and includes a custom UserForm to automate the generation of filtered reports for the operations team.

---

## 💡 Project Background & Impact

### Challenge: Automation of Report Creation
While working with the ID-NFG (Indonesia Non-Finished Goods) team, preparing the weekly LSP report was highly inefficient, taking approximately **1.5 hours** and involving repetitive manual tasks that were prone to human error.

### Solution: AI-Assisted VBA Automation
I developed an automated reporting solution by **leveraging advanced AI tools (such as Gemini/ChatGPT) for complex VBA code generation and refinement**. This approach allowed the entire process to be completed **with a single click**.

### Quantifiable Impact
| Metric | Before Automation | After Automation | Improvement |
| :--- | :--- | :--- | :--- |
| **Report Processing Time** | ~1.5 hours | **< 10 minutes** | **~90% Reduction** |
| **Manual Work** | Repetitive | Eliminated | Increased Efficiency |
| **Accuracy** | Prone to human error | Minimized Error | Increased Reliability |

---

## 🛠️ Advanced Technical Features Demonstrated

This project highlights skills in **AI integration**, data visualization, and custom application development:

### 1. **VBA Automation & Prompt Engineering**

* **AI-Assisted Development:** Demonstrates proficiency in **prompt engineering** and validating AI-generated code to create the functional `UserForm3` and all filtering logic.
* **Custom User Interface:** The "Report Generation Form" (`UserForm3`) provides a user-friendly interface for filtering data.
* **Dynamic Data Handling:** The `UserForm_Initialize` routine uses the **Scripting.Dictionary** object for efficient, fast scanning of the `LSP_Booking Sheets` worksheet to dynamically populate the filter lists (Weeks, LSPs, Dates).

### 2. **Excel Dashboard & Data Management**

* **Dynamic Charts:** Uses advanced Excel features (like Pivot Charts or data connections) to visualize weekly trends across 52 weeks of data.
* **KPI Calculation:** Uses advanced Excel formulas to calculate and display real-time KPIs (e.g., **52.8% On Time Delivery**).

---

## 📂 Repository Structure

* `dashboard/`: Contains the final, functional Excel file (`KN_Logistics_Operations_Dashboard.xlsm`).
* `vba_code/`: Contains the raw VBA code modules for review (e.g., `UserForm3.frm`).
* `data/`: Contains the simulated dataset used to drive the dashboard (`Raw_Logistics_Data.csv`).
