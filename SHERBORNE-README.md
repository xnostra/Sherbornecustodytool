# IT Asset Custody Form Automation Tool

Automated extraction of hardware specifications with direct Excel form population. Zero-dependency, fully portable, runs on any Windows computer without requiring Microsoft Excel or external modules.

## Features

✅ **Automatic Hardware Detection**
- Device type (laptop vs. desktop)
- Brand, model, and SKU
- CPU model and generation
- RAM size (rounded to standard capacities)
- Storage type and capacity
- GPU/Video controller
- Display diagonal size
- Operating system and architecture

✅ **Zero-Dependency Design**
- No external PowerShell modules required
- Uses built-in `System.IO.Compression` (available on all Windows PowerShell 5.1+)
- No Microsoft Excel installation needed
- Fully portable—copy to USB and run anywhere

✅ **Direct Excel Manipulation**
- Edits `.xlsx` file XML structure directly
- Auto-generates formatted, print-ready documents
- Automatically adjusts row heights for multi-line text
- Shrink-to-fit for long asset tag numbers
- Print-safe margins and page centering

✅ **Professional Output**
- Pre-configured print layouts
- Compliance-ready tracking forms
- Auditable asset records
- Ready to print without manual formatting

## Quick Start

### Option 1: One-Liner (Recommended)

Run this single command in PowerShell on any computer:

```powershell
irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex
```

### Option 2: Local Execution

1. Download `Fill-CustodyForm.ps1` and `custody form.xlsx` to a folder
2. Open PowerShell as Administrator
3. Run:
   ```powershell
   .\Fill-CustodyForm.ps1
   ```

### Option 3: USB Portable Deployment

Copy to USB:
```
CustodyTool/
├── Fill-CustodyForm.ps1
├── custody form.xlsx
└── Filled/          (auto-created)
```

Then run from any Windows computer—no installation needed.

## Prerequisites

- **Windows 10/11** (any edition)
- **PowerShell 5.1+** (built-in)
- **Administrator privileges** (auto-elevation if needed)
- **The custody form template** (`custody form.xlsx`)

## Usage

The script prompts for a few details after auto-detecting hardware:

```
Location (e.g., Qatar, BH, UK): Qatar
Department: IT
Staff/Custodian name [USERNAME]: John Smith
Asset Tag Number: AST-001234
```

Output is saved to: `Filled/[Name] custody [Date].xlsx`

## What Gets Auto-Detected

| Specification | Detection Method | Example |
|---------------|------------------|---------|
| Device Type | WMI Chassis | Laptop/Notebook |
| Brand & Model | Win32_ComputerSystem | Dell Latitude 5440 (SKU: 123ABC) |
| CPU | Win32_Processor | Intel Core i7-13700U (Gen 13) |
| OS | Win32_OperatingSystem | Windows 11 Pro 64-bit |
| RAM | WMI (rounded) | 16GB (actual: 15.7GB) |
| GPU | Win32_VideoController | Intel Iris Xe Graphics |
| Screen | WMI Monitor | 14.0in |
| Storage | Get-PhysicalDisk (OS drive only) | 512GB SSD |
| Serial | WMI BIOS | ABC123XYZ |

## Command Parameters

```powershell
# Use default paths (template and output in script directory)
.\Fill-CustodyForm.ps1

# Custom template path
.\Fill-CustodyForm.ps1 -TemplatePath "C:\Templates\custody-form.xlsx"

# Custom output folder
.\Fill-CustodyForm.ps1 -OutputFolder "C:\Forms\Completed"

# Both
.\Fill-CustodyForm.ps1 -TemplatePath "C:\t.xlsx" -OutputFolder "C:\out"
```

## How It Works

1. **Detection** — Queries WMI/CIM for complete hardware specs
2. **Rounding** — Converts to standard capacities (8GB, 16GB, 256GB, 512GB, etc.)
3. **Prompts** — Asks for location, department, staff name, and asset tag
4. **Generation** — Copies template, edits XML, writes values
5. **Formatting** — Applies print layouts (margins, centering, fit-to-page)
6. **Output** — Saves completed, ready-to-print Excel file

## Troubleshooting

### "Access Denied" Error
- Run PowerShell as Administrator
- If using one-liner, you may need to set execution policy first:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
  ```

### Script Downloads But Doesn't Run
- Check GitHub URL is correct
- Verify the repository is public
- Ensure internet connectivity

### Template Not Found
- Ensure `custody form.xlsx` is in the same directory as the script
- Or specify full path with `-TemplatePath` parameter

### Output Folder Permissions
- Script auto-creates `Filled\` folder
- If it fails, check NTFS permissions on the script directory

## Print Settings (Auto-Configured)

- **Orientation**: Portrait
- **Scaling**: Fit to 1 page wide × 1 page high
- **Margins**: 0.5" left/right, 0.75" top/bottom
- **Alignment**: Horizontally centered
- **Font Sizing**: Dynamic row heights for descriptions
- **Result**: Professional, print-ready form

## Version History

**v2.0** (2026-07-20)
- Complete rewrite with professional error handling
- Improved hardware detection accuracy
- Better storage type classification
- Enhanced print layout configuration
- Streamlined prompts and output

**v1.0** (Original)
- Initial release with basic functionality

## Support & Issues

For questions, issues, or suggestions:
- Open an issue on GitHub
- Check hardware detection accuracy with `Get-CimInstance` queries
- Verify Excel template structure is not corrupted

## License

Provided as-is. Modify and distribute freely.

## Author

Sherborne Custody Tool Team

---

**Last Updated**: 2026-07-20  
**Repository**: https://github.com/xnostra/Sherbornecustodytool  
**One-Liner**: `irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex`
