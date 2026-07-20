# IT Asset Custody Tracking Automation

A fully portable, zero-dependency PowerShell automation script designed to streamline the IT asset handover process. This tool automatically extracts detailed hardware specifications from the local machine and directly writes them into a formatted Excel (`.xlsx`) tracking document.

Designed for field deployments, this script manipulates the Excel file's underlying XML structure using native .NET classes, completely removing the need for Microsoft Excel, COM objects, or external PowerShell modules like `ImportExcel`.

## 🚀 Key Features

* **Zero Dependencies:** Uses built-in `System.IO.Compression` classes available in standard Windows PowerShell 5.1+. No module installation or internet access required.
* **No MS Office Required:** Directly edits the `.xlsx` file structure, making it perfect for freshly formatted machines or offline environments.
* **Automated Hardware Auditing:** Leverages WMI/CIM to accurately extract:
  * Device Category (Desktop vs. Laptop)
  * Brand, Model, and exact System SKU
  * Device Serial Number / Service Tag
  * CPU Model and Generation
  * Total usable RAM (rounded to standard marketing sizes)
  * Primary OS Drive Storage (ignoring USBs, mapped to standard capacities)
  * Display Diagonal Size (converted to inches)
  * GPU / Video Controller details
* **Smart Excel Formatting:** Dynamically adjusts row heights for multi-line text, shrinks long text to fit specific cells, and injects print-safe page layouts (margins, centering, fit-to-page).
* **Auto-Elevation:** Automatically detects permission issues and requests Admin elevation if restricted directory access is encountered.

## 📋 Folder Structure

For the script to function correctly, it must be run alongside the target Excel template. The recommended layout for a portable USB drive is:

```text
CustodyTool/
├── Fill-CustodyForm.ps1       # The main execution script
├── custody form.xlsx          # The blank template (MUST remain next to the script)
└── Filled/                    # Auto-generated directory for completed forms
