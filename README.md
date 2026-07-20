# Sherborne Custody Tool

A PowerShell automation tool for efficiently managing and populating Sherborne custody forms with comprehensive data management capabilities.

## Overview

The Sherborne Custody Tool streamlines the process of handling custody form data through automated population and validation. This tool is designed to reduce manual data entry, minimize errors, and improve workflow efficiency for custody form management operations.

## Features

- **Automated Form Population** — Automatically populate custody forms from structured data sources
- - **Data Validation** — Built-in validation to ensure data accuracy and integrity
  - - **Batch Processing** — Process multiple forms efficiently in batch operations
    - - **Error Handling** — Comprehensive error reporting and logging capabilities
      - - **Excel Integration** — Seamless integration with Excel spreadsheet data (custody form.xlsx)
        - - **PowerShell Scripting** — Full PowerShell automation capabilities for custom workflows
         
          - ## Getting Started
         
          - ### Prerequisites
         
          - - Windows operating system with PowerShell 5.0 or higher
            - - Microsoft Excel (for working with custody form templates)
              - - Appropriate file permissions for target directories
               
                - ### Installation
               
                - 1. Clone this repository to your local machine
                  2. 2. Extract the files to your working directory
                     3. 3. Run the Fill-CustodyForm.ps1 script with appropriate parameters
                       
                        4. ### Usage
                       
                        5. ```powershell
                           .\Fill-CustodyForm.ps1 -InputFile "custody form.xlsx" -OutputPath ".\output"
                           ```

                           Refer to the script comments for detailed parameter descriptions and usage examples.

                           ## Files

                           - **Fill-CustodyForm.ps1** — Main PowerShell automation script
                           - - **custody form.xlsx** — Template and sample custody form data
                             - - **.gitattributes** — Git configuration for file handling
                              
                               - ## Contributing
                              
                               - Contributions are welcome. Please ensure any modifications maintain code quality and documentation standards.
                              
                               - ## License
                              
                               - This project is provided as-is for authorized use.
                              
                               - ## Support
                              
                               - For issues, questions, or feature requests, please contact the repository maintainer.
