# SAP Expense Confirmation Automation - UiPath RPA

This RPA project automates the process of confirming expenses in SAP. It includes handling multiple scenarios such as Purchase Order (PO) confirmation, Inscoped expenses, and Expense Import. The automation is designed to run based on specific date rules and excludes weekends.

## Project Description

The automation retrieves expense data, validates it, and processes it into SAP via UI automation. The process includes two operational conditions:

- **W1 (Week 1 Process)**  
  - Runs on the **15th, 16th, and 17th of each month** (excluding weekends).  
  - Handles the first cycle of expense confirmation.

- **W2 (Week 2 Process)**  
  - Runs on the **1st, 2nd, and 3rd of each month** (excluding weekends).  
  - Handles the second cycle of expense confirmation.

The automation will automatically determine whether to execute the W1 or W2 process based on the current date.

## Features

- Date-based execution:
  - **W1 → 15, 16, 17**
  - **W2 → 1, 2, 3**
  - Skips weekends automatically
- Automated SAP processing for:
  - Purchase Order (PO) Expense
  - Inscoped Expense
  - Expense Import
- Generates Excel output reports
- Modular workflow design
- Exception handling with detailed logs and screenshots

## Project Structure

| Folder/File                 | Description                                                   |
|-----------------------------|---------------------------------------------------------------|
| Main.xaml                   | Main entry point for the SAP Expense Confirmation process     |
| Modular/                    | Reusable workflows for SAP PO, Inscoped, and Import processes |
| Framework/                  | REFramework components                                        |
| Data/Config.xlsx             | Configuration file (SAP, folder paths, etc.)                 |
| Data/Output/                | Output reports                                                |
| Screenshots/                | Sample screenshots                                            |
| Exceptions_Screenshots/     | Error screenshots                                             |
| project.json                 | UiPath project metadata                                       |
| README.md                    | Project documentation                                         |

## Process Workflow

1. **Date Check:**  
   - The robot checks today's date.  
   - If today is **15, 16, or 17** → Runs **W1 Process** (if not weekend).  
   - If today is **1, 2, or 3** → Runs **W2 Process** (if not weekend).  
   - Otherwise, the process does not run.

2. **Retrieve Expense Data**  
3. **Validate Data**  
4. **Process in SAP**  
   - Expense Import
   - PO Confirmation
   - Inscoped Expense

5. **Generate Output Report**  
6. **Exception Handling with Logs and Screenshots**  

## How to Run

1. Open the project in UiPath Studio.
2. Configure `Data/Config.xlsx`:
   - SAP credentials
   - Folder paths
3. Run `Main.xaml` manually or deploy via Orchestrator.  
   - The bot will determine whether it's W1 or W2 based on the current date.

## Exception Handling

- Logs error with screenshot in `Exceptions_Screenshots`.
- Generates error reports for review.

## Contact

- Email: fadillah650@gmail.com
- LinkedIn: https://linkedin.com/in/enrico-naufal-fadilla-54338a256
