# IT Asset Custody Form Automation

Automated extraction of hardware specifications with direct Excel form population. Zero-dependency, fully portable, runs on any Windows computer without requiring Microsoft Excel or external modules.

## Quick Start

Run this single command in PowerShell on any computer:

```powershell
irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex
```

That's it—the one-liner auto-downloads the script and Excel template, auto-elevates to admin, auto-detects hardware, generates the completed form, **saves it to your Desktop**, and opens it automatically.

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

## Prerequisites

- Windows 10/11 (any edition)
- PowerShell 5.1+ (built-in)
- Administrator privileges (auto-elevated if needed)
- The custody form template (`custody form.xlsx`)

## Usage

The script prompts for details after detecting hardware:

```
Location (e.g., Qatar, BH, UK): Qatar
Department: IT
Staff/Custodian name [USERNAME]: John Smith
Asset Tag Number: AST-001234
```

Output filename: `[Name] custody [Date].xlsx`
- Via the one-liner: saved to your **Desktop** and opened automatically
- Via local execution: saved to the `Filled/` folder next to the script

## Deployment Options

**Option 1: One-Liner** (Recommended)

```powershell
irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex
```

**Option 2: Local Execution**

```powershell
.\Fill-CustodyForm.ps1
```

**Option 3: USB Portable**

Copy to USB:
```
CustodyTool/
├── Fill-CustodyForm.ps1
├── custody form.xlsx
└── Filled/
```

Then run from any Windows computer.

## Command Parameters

```powershell
# Default
.\Fill-CustodyForm.ps1

# Custom template
.\Fill-CustodyForm.ps1 -TemplatePath "C:\Templates\form.xlsx"

# Custom output
.\Fill-CustodyForm.ps1 -OutputFolder "C:\Output"

# Both
.\Fill-CustodyForm.ps1 -TemplatePath "C:\t.xlsx" -OutputFolder "C:\out"
```

## Auto-Detected Specifications

Device type, brand, model, SKU, CPU with generation, OS, RAM (rounded to standard sizes), GPU, screen diagonal, storage capacity and type (OS drive only), and device serial number.

## Troubleshooting

**Access Denied**
- Run PowerShell as Administrator
- Or: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force`

**Template Not Found**
- Ensure `custody form.xlsx` is in script directory
- Or specify full path with `-TemplatePath` parameter

**Download Fails**
- Verify GitHub URL is correct
- Check internet connectivity

## Print Configuration

Auto-configured:
- Portrait orientation
- Fit to 1 page wide (fills the A4 width, no wasted whitespace)
- 0.5" left/right margins
- Horizontally centered
- Dynamic row heights

## Version History

**v2.1** (2026-07-20)
- One-liner now auto-downloads the template and saves output to the Desktop
- Output form auto-opens when complete
- Print layout fits A4 width (removed excess whitespace)

**v2.0** (2026-07-20)
- Professional refactor with error handling
- One-liner launcher support
- Improved hardware detection

**v1.0**
- Initial release

## Support

For issues: Open a GitHub issue or check hardware detection with `Get-CimInstance` queries.

## License

Provided as-is. Modify and distribute freely.

---

**One-Liner**: `irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex`

**Repository**: https://github.com/xnostra/Sherbornecustodytool

**Last Updated**: 2026-07-20
