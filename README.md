# KN-Logistics-Report-Automation-Project
VBA automation solution created during a Data Operations internship at Kuehne + Nagel to efficiently generate filtered logistics reports from operational data

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
