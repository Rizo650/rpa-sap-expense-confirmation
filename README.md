# SAP Expense Confirmation - UiPath RPA

This RPA project automates the SAP Expense Confirmation process using UiPath. It integrates Excel VBA macros for complex report manipulation and performs automated SAP data input based on dynamic date rules. The solution is designed for reliability, scalability, and accuracy in handling financial data.

---

## Project Description

This automation is designed to process expense confirmations in SAP based on two operational cycles:

- **W1 Process:**  
  Runs on **15th, 16th, and 17th** of each month, excluding weekends.

- **W2 Process:**  
  Runs on **1st, 2nd, and 3rd** of each month, excluding weekends.

The process includes retrieving raw data, preparing Excel reports using VBA, validating the data, and then processing it into SAP. Output reports are generated for record-keeping and are later used for email notifications.

---

## Key Features

- **Date-driven execution:**
  - **W1:** Runs on 15, 16, 17 (skip weekends)
  - **W2:** Runs on 1, 2, 3 (skip weekends)
- **VBA-powered Excel manipulation:**
  - Cleans raw data
  - Formats reports
  - Transforms datasets for SAP input
- Automated SAP processing for:
  - Purchase Order (PO) Expenses
  - Inscoped Expenses
  - Expense Import
- Output Excel report generation
- Modular workflow design using REFramework
- Robust exception handling with logs and screenshots

---

## Excel Processing with VBA

This project relies heavily on VBA (Visual Basic for Applications) scripts to perform complex manipulations on Excel files that are impractical or inefficient with native UiPath Excel activities.

### VBA Usage Includes:
- Data cleaning and validation
- Complex formula application
- Removing unnecessary columns/rows
- Sheet formatting
- Merging or splitting sheets
- Generating pivot tables or summaries
- Preparing structured reports for SAP input

### Integration Method:
- VBA is either:
  - Embedded within Excel templates (`.xlsm` files)
  - Or stored externally in `/Data/VBA/` folder
- Executed using **UiPath’s "Execute Macro"** activity.

### VBA Requirements:
- Excel installed with macros enabled
- Trust access to VBA project object model activated

---

## Project Structure

| Folder/File                 | Description                                                   |
|-----------------------------|---------------------------------------------------------------|
| Main.xaml                   | Main entry point for SAP Expense Confirmation                 |
| Modular/                    | Sub-workflows (e.g., PO, Inscoped, Import, VBA executors)     |
| Framework/                  | REFramework components                                        |
| Data/Config.xlsx             | Configuration (SAP credentials, folder paths, parameters)     |
| Data/Output/                | Generated output reports                                      |
| Data/VBA/                   | VBA macro files (if external)                                 |
| Screenshots/                | Sample screenshots of successful process steps                |
| Exceptions_Screenshots/     | Error screenshots captured during failure                     |
| project.json                 | UiPath project metadata                                       |
| README.md                    | Project documentation                                         |

---

## Process Workflow

### 1. **Date Check Logic**
- Check today's date:
  - If **15, 16, 17** → Execute **W1 process** *(if not weekend)*
  - If **1, 2, 3** → Execute **W2 process** *(if not weekend)*
  - If not matching → Skip process or end gracefully

### 2. **Retrieve Raw Data**
- Pull data from SAP export.

### 3. **Excel Report Preparation (with VBA)**
- Execute VBA macros to:
  - Clean raw data
  - Apply necessary formatting
  - Transform data into SAP input-ready format

### 4. **Data Validation**
- Check for missing, invalid, or duplicate data.

### 5. **SAP Processing**
- Use UI automation to:
  - Input data into SAP for:
    - Purchase Orders (PO)
    - Inscoped expenses
    - Expense imports

### 6. **Generate Output Report**
- Save results into `/Data/Output/` with naming convention per date and process type (W1/W2).

### 7. **Exception Handling**
- Capture any errors with screenshots
- Log detailed error messages
- Store screenshots in `Exceptions_Screenshots`

---

## How to Run

1. Open the project in **UiPath Studio**.
2. Configure `Data/Config.xlsx`:
   - SAP credentials
   - Folder paths for input/output
   - VBA macro file locations (if external)
   - Other runtime parameters
3. Run `Main.xaml` directly or deploy via **UiPath Orchestrator**.
4. The bot will determine whether to execute the **W1** or **W2** process based on the current date.

---

## Exception Handling

- Logs all errors with detailed messages.
- Captures UI screenshots during exceptions.
- Screenshots stored in `Exceptions_Screenshots` folder.
- Error reports can be used for manual review or escalation.

---

## Requirements

- UiPath Studio (Enterprise)
- Microsoft Excel with macros enabled
- SAP GUI installed and configured for automation
- Access rights to SAP transaction screens involved in the process

---

## Contact

For questions, improvements, or collaboration:

- **Email:** fadillah650@gmail.com  
- **LinkedIn:** [Enrico Naufal Fadilla](https://linkedin.com/in/enrico-naufal-fadilla-54338a256)
