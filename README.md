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
- Via local execution or the `.bat`: saved to the `Filled/` folder (or Desktop, if you passed `-OutputFolder`) and opened automatically
- Optionally, auto-emailed as an attachment if email relay is configured (see Email Automation below)

## Deployment Options

**Option 1: One-Liner** (Recommended)

```powershell
irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex
```

**Option 2: Double-click launcher (works even if "running scripts is disabled")**

Double-click **`Run Custody Form.bat`** in the folder. This runs the tool without ever executing `Fill-CustodyForm.ps1` as a script file, so it works even on a freshly-imaged or locked-down computer where Windows' Execution Policy blocks running `.ps1` files outright — no setup, no admin needed just to launch it.

**Option 3: Local Execution (PowerShell prompt)**

```powershell
.\Fill-CustodyForm.ps1
```

**Option 4: USB Portable**

Copy to USB:
```
CustodyTool/
├── Fill-CustodyForm.ps1
├── Run Custody Form.bat
├── custody form.xlsx
└── Filled/
```

Then run `Run Custody Form.bat` from any Windows computer.

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

## Email Automation (optional)

The tool can automatically email the finished form as an attachment to `jcarlos@sherborneqatar.org` - useful on freshly-imaged computers with no Outlook profile set up. It does **not** use a stored password, Outlook, or any secret in this repo: the script uploads the finished form into a OneDrive folder (signing in with a one-time device-code prompt - the same "go to microsoft.com/devicelogin and enter this code" flow used by many Microsoft CLI tools), and a Power Automate flow watching that folder sends the email automatically using your own Microsoft 365 connection.

**One-time setup (Power Automate, ~10 minutes, uses the free tier - no Premium license needed):**

1. In [OneDrive](https://onedrive.com) (signed in with your Microsoft 365 account), create a folder named exactly **`CustodyFormsToEmail`**.
2. Go to [make.powerautomate.com](https://make.powerautomate.com) and create a new **Automated cloud flow**.
3. For the trigger, search for and select **"When a file is created"** under the **OneDrive for Business** connector (not the SharePoint one, and not the HTTP trigger - both of those either don't fit or need a Premium license).
4. Set its **Folder** to `CustodyFormsToEmail`.
5. Add an action: **"Send an email (V2)"** (Office 365 Outlook) with:
   - To: `jcarlos@sherborneqatar.org` (typed directly)
   - Subject: dynamic content → **File name**
   - Body: anything, e.g. "A custody form was submitted. See attached."
   - Under **Show advanced options** → **Attachments**: Attachment Name → **File name** (dynamic content), Attachment Content → **File content** (dynamic content - no expression needed, OneDrive provides it directly)
6. Save the flow.

That's it - no URL to copy into the scripts. Once the flow is saved, every run of the tool (with emailing left on - the default) will prompt you once to sign in via a device code, then upload the finished form to that OneDrive folder, which triggers the flow and emails it automatically.

**The email isn't instant** - the free OneDrive trigger checks the folder periodically rather than reacting immediately, so it can take a few minutes for the email to arrive after the upload succeeds. That delay is normal, not a sign anything's broken.

**Signing in:** the tool opens your browser to Microsoft's sign-in page and copies the code to your clipboard automatically - just paste it into the box on that page (Microsoft's device sign-in page does not support auto-filling the code; typing/pasting it in is a one-time Microsoft limitation, not something this tool can skip).

**To turn auto-email off**, either:
- In `Run Custody Form.bat`: change `set "AUTO_EMAIL=1"` to `set "AUTO_EMAIL=0"`, or
- In `invoke-custody.ps1`: change `$autoEmail = $true` to `$autoEmail = $false`

If signing in or uploading fails for any reason (no internet, sign-in skipped, OneDrive unreachable), the tool still saves and opens the file normally - it just skips the email with a warning, and full error details are saved to `CustodyForm.log` next to the script (Notepad opens automatically with the log whenever something goes wrong, so you can read/copy the exact error even after the window closes).

**Why this design instead of a stored password?** This repository is public on GitHub. Any password or API key placed directly in a public script is visible to anyone who views it (even in old commit history) and could be used to send mail as your organization. Device-code sign-in never exposes a password - it only ever produces a short-lived token for whoever is sitting at the keyboard, requested fresh every run.

## Auto-Detected Specifications

Device type, brand, model, SKU, CPU with generation, OS, RAM (rounded to standard sizes), GPU, screen diagonal, storage capacity and type (OS drive only), and device serial number.

## Troubleshooting

**"Running scripts is disabled on this system"**
- Use **`Run Custody Form.bat`** instead of running the `.ps1` directly - it bypasses this entirely, including on computers where the policy is locked by group policy (see Deployment Options above)
- The tool also auto-relaxes the policy for your own user account on every run, but that can't override a policy locked by IT/group policy - the `.bat` is the reliable fix

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

**v2.6** (2026-07-22)
- Fixed a real bug: emailing could fail with "file in use" because the form was already open in Excel by the time the upload tried to read it. The form's contents are now read into memory before it's opened in Excel, so emailing no longer depends on the file being closed.

**v2.5** (2026-07-22)
- Every run now writes a `CustodyForm.log` file next to the script - if something goes wrong, Notepad opens automatically with the full error detail so it's easy to read and copy, even after the tool's window closes
- The sign-in code is now copied to your clipboard automatically, and the messaging no longer implies the code auto-fills on Microsoft's page (it doesn't - that's a Microsoft limitation)

**v2.4** (2026-07-22)
- Email automation reworked to use a free OneDrive-triggered Power Automate flow instead of an HTTP trigger (the HTTP trigger requires a Premium license many tenants don't have) - now uses a one-time device-code sign-in to upload the form instead

**v2.3** (2026-07-22)
- Row 9 height is now fixed at 172.8 (was auto-calculated from description line count)
- The form now auto-opens no matter which launcher is used (previously only the one-liner opened it - direct/`.bat` runs did not)
- Added optional auto-email of the finished form via a private relay (see Email Automation above) - no credentials stored in this repo

**v2.2** (2026-07-22)
- Added `Run Custody Form.bat` - a double-click launcher that runs on any computer, even ones where "running scripts is disabled" by policy (no more Execution Policy errors)
- The one-liner and the admin-elevation relaunch no longer call the script as a file internally, so both keep working on locked-down machines too

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

MIT — see [`LICENSE`](LICENSE).

---

**One-Liner**: `irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex`

**Repository**: https://github.com/xnostra/Sherbornecustodytool
