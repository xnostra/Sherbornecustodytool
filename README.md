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

✅ **On-Screen Signature**
- Sign right in the popup with a mouse or trackpad - no printing, signing on paper, and re-scanning
- Embedded directly into the form's existing "Signature" box
- Optional - leave it blank and the form generates the same as before

## Prerequisites

- Windows 10/11 (any edition)
- PowerShell 5.1+ (built-in)
- Administrator privileges (auto-elevated if needed)
- The custody form template (`custody form.xlsx`)

## Usage

First, a popup asks **"Is this a computer or laptop?"**
- **Yes** - the tool detects hardware automatically (model, CPU, RAM, storage, etc.), same as before.
- **No** - for items you can't run this tool on yourself (a projector, clicker, screen, etc.), typically filled out from your own computer. Hardware auto-detection is skipped, and the popup form below includes two extra boxes so you can type in the item's category and description by hand.

Then a popup window asks for the remaining details:
- **Location** - a dropdown, choose one of: Mall of Qatar, Bani Hajer, Boys School, Girls School
- **Department** - type it in
- **Custodian name** - type it in, or leave blank to use your own Windows username
- **Asset Tag Number** - type it in
- **Signature** - a small white box where you can draw your signature with a mouse or trackpad (click and drag). A "Clear signature" button lets you redo it if needed. This is optional - leaving it blank just leaves the form's Signature box empty, same as before this feature existed.
- *(only if you chose "No" above)* **Item Category** and **Item Description** - type them in

All of this gets written onto the form in CAPS automatically (the signature is inserted as-drawn, not capitalized). The form won't let you continue until the required boxes are filled in - the signature is the only optional one.

Output filename: `[Name] custody [Date].xlsx`
- Saved into a subfolder named after the **Location** you picked (e.g. `Filled\MALL OF QATAR\Name custody Date.xlsx`), purely to keep forms from different sites organized instead of all piling up loose in one folder
- Via the one-liner: saved under a location subfolder on your **Desktop** and opened automatically
- Via local execution or the `.bat`: saved under a location subfolder inside `Filled/` (or under `-OutputFolder`, if passed) and opened automatically
- Optionally, auto-emailed as an attachment if email relay is configured (see Email Automation below) - the OneDrive copy used for emailing is organized into a matching location subfolder too

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

The tool can automatically email the finished form as an attachment - useful on freshly-imaged computers with no Outlook profile set up. It does **not** use a stored password, Outlook, or any secret in this repo: the script uploads the finished form into a OneDrive folder (signing in with a one-time device-code prompt - the same "go to microsoft.com/devicelogin and enter this code" flow used by many Microsoft CLI tools), and a Power Automate flow watching that folder sends the email automatically using your own Microsoft 365 connection.

When emailing is on, a popup lets whoever is running the tool pick **"Default (jcarlos@sherborneqatar.org)"** or **"IT Group (itn@sherborneqatar.org)"** as the destination for that particular form - e.g. your manager (or anyone else) can choose to send a copy to the group inbox instead of you personally. The tool also automatically CCs whoever actually signed in and uploaded the form, so the person who ran it always gets a copy too - the email itself is still sent via jcarlos's own Outlook connection (Power Automate connections are fixed per-flow, not per-run), but the To/CC lines reflect the right people.

**How routing works under the hood:** the "which inbox" (`[ITN]`) and "who to CC" (`[CC someone_AT_domain.com]`) info travels as tags inside the **filename itself** of the copy uploaded to OneDrive, e.g. `[ITN] [CC someone_AT_domain.com] Name custody Date.xlsx`. The flow reads these tags straight off the trigger's filename output to build the Subject, CC, and Attachment Name - each of which strips the tags back off before the recipient ever sees them, so the email itself always shows a clean subject/attachment name no matter what the OneDrive copy is named.

The OneDrive copy is also placed into a **location subfolder** purely for organization (e.g. `CustodyFormsToEmail/MALL OF QATAR/[CC ...] Name custody Date.xlsx`) - this subfolder has no effect on routing, it's just tidiness so forms from different sites don't all pile up in one place. Since a subfolder is involved, the trigger has **Include subfolders** turned on (see setup step 4 below); the flow still gets the routing tags from the filename, not the folder.

**One-time setup (Power Automate, ~10 minutes, uses the free tier - no Premium license needed):**

1. In [OneDrive](https://onedrive.com) (signed in with your Microsoft 365 account), create a folder named exactly **`CustodyFormsToEmail`**.
2. Go to [make.powerautomate.com](https://make.powerautomate.com) and create a new **Automated cloud flow**.
3. For the trigger, search for and select **"When a file is created"** under the **OneDrive for Business** connector (not the SharePoint one, and not the HTTP trigger - both of those either don't fit or need a Premium license).
4. Set its **Folder** to `CustodyFormsToEmail`, then open **Advanced parameters** and set **Include subfolders** to **Yes** - required so the trigger sees files uploaded into a location subfolder (e.g. `CustodyFormsToEmail/MALL OF QATAR/...`) rather than only files sitting directly in `CustodyFormsToEmail` itself.
5. Add **"Send an email (V2)"** (Office 365 Outlook) with To: `jcarlos@sherborneqatar.org`, and these field mappings (all under the **Expression** tab):
   - **Subject** →
     `if(startsWith(if(startsWith(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), '[ITN] '), substring(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), 6), base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded'])), '[CC '), substring(if(startsWith(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), '[ITN] '), substring(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), 6), base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded'])), add(indexOf(if(startsWith(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), '[ITN] '), substring(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), 6), base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded'])), '] '), 2)), if(startsWith(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), '[ITN] '), substring(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), 6), base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded'])))`
     This strips a leading `[ITN] ` tag if present, then strips a `[CC ...] ` tag if present, leaving just the plain filename.
   - Body → anything, e.g. "A custody form was submitted. See attached."
   - **CC** →
     `if(contains(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), '[CC '), replace(substring(substring(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), add(indexOf(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), '[CC '), 4)), 0, indexOf(substring(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), add(indexOf(base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded']), '[CC '), 4)), ']')), '_AT_', '@'), '')`
     Pulls out whatever's between `[CC ` and the closing `]`, swaps `_AT_` back to `@`. If there's no CC tag, this safely returns an empty string (no CC added).
   - Under **Show advanced options** → **Attachments**: Attachment Name → same expression as Subject above, Attachment Content → **File content** (dynamic content).
6. Save the flow.

That's it - no URL to copy into the scripts. Once the flow is saved, every run of the tool (with emailing left on - the default) will prompt whoever's running it to sign in via a device code, ask which inbox this form should go to, then upload the finished form (tagged filename, location subfolder) to OneDrive, which triggers the flow and CCs the uploader automatically.

**Note on the "IT Group" inbox choice:** the live flow now has a **Condition** action right after the trigger, checking whether `base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded'])` **starts with** `[ITN] `. The **True** branch has its own "Send an email (V2)" action with To: `itn@sherborneqatar.org` (a distribution group, not an individual mailbox - works exactly like any other address in a "To" field); the **False** branch has its own "Send an email (V2)" action with To: `jcarlos@sherborneqatar.org`. Both branches use the same Subject/CC/Attachment expressions above. So picking "IT Group" in the popup now actually routes the email to itn@sherborneqatar.org, not just tags the filename.

If you'd rather keep it simpler and always send to one fixed address (no CC parsing at all), you can just map Subject/Attachment Name straight to `base64ToString(triggerOutputs()?['headers/x-ms-file-name-encoded'])` and leave CC blank - the email's subject/attachment name will just show the raw tagged filename instead of a cleaned-up one.

**The email isn't instant** - the free OneDrive trigger checks the folder periodically rather than reacting immediately, so it can take a few minutes for the email to arrive after the upload succeeds. That delay is normal, not a sign anything's broken.

**Signing in:** the tool opens the browser to Microsoft's sign-in page, copies the code to the clipboard, and (best-effort) auto-pastes it into the code box on that page for you, so there's normally nothing to type before your password. If the browser window can't be auto-focused for any reason, the code is still on your clipboard - just paste it in yourself as before. **Your password always stays 100% manual** - this automation only ever touches the code box; it never sees, enters, or interacts with the password field or anything after it, by design and with no exception.

**Letting other staff run this tool (not just jcarlos):** uploads always target jcarlos's OneDrive folder specifically (via a hardcoded Drive ID in the script), regardless of which Microsoft 365 account signs in during the device-code step. This means:
- Any staff member can run the tool on any computer, sign in with their **own** Microsoft 365 account, and it still lands in jcarlos's `CustodyFormsToEmail` folder and emails jcarlos (or itn, plus a CC to themselves) - no per-person setup needed on their end.
- For this to work, `CustodyFormsToEmail` needs org-wide edit access granted directly on the folder - **this is not the same as a "sharing link"**. In OneDrive, right-click the folder → look for **Manage access** (or open the **Details/Info** panel → **Manage access**) → **Grant access** → type your organization's name (e.g. "Sherborne Qatar School") and select the org-wide group entry that appears → set it to **Can edit** → confirm. A "Share → Anyone with the link" or "Share → People in [organization]" **link** does *not* grant this - links and direct folder permissions are two different things in OneDrive, and only the direct permission (via Manage Access) is what the script's Graph API calls actually check.
- If the folder's permissions are ever changed to something more restrictive, other staff's sign-ins will fail to upload with a `403 Forbidden` error (their own device-code sign-in still works, but the write to jcarlos's folder will be denied) - the tool will just skip emailing with a warning in that case, everything else still works normally.

**Why the file in OneDrive briefly shows a tagged name like `[CC someone_AT_domain.com] Name custody Date.xlsx`:** the free-tier OneDrive trigger has no way to read custom metadata, so the `[CC ...]`/`[ITN]` routing info has to travel as part of the filename itself at upload time - there's no other channel available to pass it along. About 8 minutes after upload, a background process (launched by the tool, but running independently of it - closing the tool's window right after upload does not stop it) renames the OneDrive copy back to the plain, clean filename (e.g. `Name custody Date.xlsx`), so the reference copy sitting in `CustodyFormsToEmail` ends up tidy long-term. The email itself never shows the tag either way (see Subject/CC expressions above) - this rename is purely about how the file looks if you browse to it in OneDrive afterward.

**Residual risk of the 8-minute rename (accepted tradeoff):** the free-tier trigger's polling interval isn't fixed, predictable, or controllable by this script. Live testing (checked against actual Power Automate run history) showed the trigger firing anywhere from about 1 to 5+ minutes after upload - the older 90-second delay was regularly losing this race and silently dropping the CC/ITN tag on every run before it was caught and fixed in v3.9. 8 minutes is comfortably past everything observed so far, but it's still not a hard guarantee - on an unusually slow poll cycle it's theoretically possible for the rename to still happen before the trigger has read the file, which would silently drop the CC/ITN tag the same way the old 90-second (and earlier, 5-second) versions could. This tradeoff (a clean OneDrive filename, with a small residual risk, instead of a fully risk-free but more complex "confirm the flow actually fired first" check) was discussed and deliberately accepted. If a CC ever appears to go missing, check `CustodyForm.log` for the upload/rename timing on that run.

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

**Signature box doesn't seem to draw anything**
- Make sure you're clicking and holding the left mouse button (or pressing down on the trackpad) while moving, not just hovering
- Use "Clear signature" and try again if a stray mark gets left behind

**Sign-in code didn't auto-paste**
- This is best-effort - it depends on being able to find and focus the newly-opened browser window, which can occasionally fail (e.g. a different window stole focus first)
- The code is still copied to your clipboard regardless - just click into the code box on the Microsoft sign-in page and paste it yourself (Ctrl+V)

## Print Configuration

Auto-configured:
- Portrait orientation
- Fit to 1 page wide (fills the A4 width, no wasted whitespace)
- 0.5" left/right margins
- Horizontally centered
- Dynamic row heights

## Version History

**v3.9** (2026-07-23)
- Fixed a real bug, confirmed live via Power Automate run history: the automatic CC (the person who signed in and uploaded the form) was coming back empty on every recent run. Root cause was the OneDrive filename rename (which strips the `[CC ...]` tag for tidiness) consistently happening *before* the free-tier trigger got around to reading the file - the trigger was observed firing anywhere from about 1 to 5+ minutes after upload, well past the old 90-second delay, so the tag was reliably gone by the time the flow read the filename. The rename delay is now 8 minutes (comfortably past everything observed, though still not a hard guarantee - see "Residual risk" note under Email Automation), and it now runs in a separate background process instead of blocking the tool's console window, so there's no need to sit and wait for it - the window can be closed right after upload and the rename still happens invisibly a few minutes later.

**v3.8** (2026-07-23)
- Fixed a real bug: choosing "IT Group (itn@sherborneqatar.org)" in the popup tagged the filename correctly, but the live Power Automate flow had no branching logic - it always sent to jcarlos@sherborneqatar.org regardless of which option was picked. The flow now has a Condition action that checks the `[ITN] ` filename tag and sends to itn@sherborneqatar.org (True branch) or jcarlos@sherborneqatar.org (False branch) accordingly - both branches keep the same Subject/CC/Attachment expressions as before. No script changes were needed - the filename tagging this depends on was already correct.

**v3.7** (2026-07-23)
- Reverted the v3.6 subfolder-based CC/ITN routing - it worked correctly but created a routing subfolder inside `CustodyFormsToEmail` in OneDrive that wasn't wanted, purely for aesthetic/organizational reasons. Routing info is back to living in the uploaded file's name (as in v3.5), and the corresponding Power Automate Compose steps (`Parse Routing Tag`, `Parse CC Raw`) were deleted from the live flow.
- To still get a clean-looking filename in OneDrive without the subfolder, the tool now renames the OneDrive copy back to a clean name **90 seconds** after upload (much longer than the unsafe 5-second version from v3.4, which raced the trigger and could silently lose the CC). This is a fixed delay, not a guarantee - see "Residual risk" note under Email Automation above. This tradeoff was discussed and deliberately accepted in exchange for a simpler, subfolder-free OneDrive layout.
- New: saved forms (both the local Desktop/`Filled` copy and the OneDrive copy used for emailing) are now automatically organized into a subfolder named after the **Location** you pick in the popup (e.g. `Filled\MALL OF QATAR\...` and `CustodyFormsToEmail\MALL OF QATAR\...`), purely for tidiness - this has no effect on CC/ITN routing, which still comes from the filename tags. Requires the Power Automate trigger's **Include subfolders** to stay set to **Yes** (same setting as v3.6, now serving a different purpose).

**v3.6** (2026-07-23)
- The OneDrive filename is now clean and readable from the moment it's uploaded (e.g. `Name custody Date.xlsx`), instead of permanently carrying the `[CC ...]`/`[ITN]` routing tag (as in v3.5 and earlier). The routing info moved to a **subfolder name** instead (e.g. `CustodyFormsToEmail/CC=someone_AT_domain.com,ITN/Name custody Date.xlsx`) - this achieves the clean-filename goal from v3.4 without its bug, since the filename itself never needs to change after upload (no race with the flow's trigger). Requires a one-time Power Automate update: turn on "Include subfolders" on the trigger, add two Compose steps to parse the folder tag, and simplify the CC/Subject/Attachment Name expressions (see Email Automation above for exact steps). Existing single-recipient flows that don't add these steps will keep working but won't get a clean filename or CC.

**v3.5** (2026-07-23)
- Reverted the v3.4 change (see below) - it caused a real bug: renaming the OneDrive copy back to a clean name a few seconds after upload sometimes happened *before* the Power Automate trigger had actually read the tagged filename, which silently dropped the CC entirely (confirmed live: a run's CC came back empty because the trigger read the already-renamed clean name). The free-tier trigger's polling interval isn't something this tool can know or control, so there's no delay that's reliably safe. The OneDrive copy now permanently keeps its tagged `[CC ...]`/`[ITN]` filename again, exactly like v3.3 and earlier - correct CC matters more than a tidy filename in OneDrive. The email itself was never affected by any of this either way.

**v3.4** (2026-07-23) - reverted in v3.5, see above
- The OneDrive copy used for emailing was renamed back to a clean filename a few seconds after upload. This turned out to be unsafe (see v3.5) and was reverted.

**v3.3** (2026-07-23)
- Added an on-screen signature box to the popup form - draw your signature with a mouse or trackpad and it's embedded directly into the form's existing "Signature" cell. Optional - leave it blank and the form generates exactly as before.
- The Microsoft sign-in code is now auto-pasted into the browser's code box (best-effort - falls back to the clipboard-copy-and-paste-it-yourself behavior if the browser window can't be auto-focused). Password entry is untouched and always stays fully manual - this automation only ever interacts with the code box, nothing after it.

**v3.2** (2026-07-23)
- The tool now automatically CCs whoever actually signed in and uploaded the form on the notification email, so the person who ran it always gets their own copy - regardless of whether it went to jcarlos or the IT Group. Requires a small addition to the Power Automate flow (a CC expression - see Email Automation above); without that addition, emails keep sending exactly as before, just without the CC.

**v3.1** (2026-07-22)
- Auto-email now supports two destinations: when emailing is on, a popup lets you choose "Default (jcarlos@sherborneqatar.org)" or "IT Group (itn@sherborneqatar.org)" for that form. Requires a one-time update to the Power Automate flow (see Email Automation above) to add a Condition step that routes based on the choice - existing single-recipient flows keep working as before until that's added.

**v3.0** (2026-07-22) - Production-readiness pass
- Fixed a real data-loss risk: two custody forms generated for the same person on the same day (e.g. two different assets, or a shared "Filled" folder) used to silently overwrite each other. The tool now automatically adds "(2)", "(3)", etc. to the filename so nothing is ever silently overwritten.
- `CustodyForm.log` no longer grows forever - it's automatically trimmed once it passes ~2MB, keeping the most recent history instead of accumulating indefinitely.
- Added a quick internet-connectivity check before attempting to email the form - on a computer with no internet (common right after imaging, before it's joined to Wi-Fi), the tool now skips straight to a clear "no internet, skipping email" message instead of waiting through a slower timeout.
- Fixed a bug where requesting admin elevation (on a locked-down account) would silently drop the `-EmailForm` request - a run started with auto-email on now keeps that setting through an elevation prompt instead of finishing without emailing.

**v2.9** (2026-07-22)
- Smarter model name detection now works well across brands, not just Lenovo - Dell, HP, Microsoft Surface, ASUS, Acer, and others now show a clean "Brand Model" name (e.g. "Dell Latitude 5420", "HP EliteBook 840 G8") instead of a raw/duplicated string, with brand names normalized to how they're commonly known (e.g. "Dell Inc." -> "Dell", "ASUSTeK COMPUTER INC." -> "ASUS"). Generic/white-box machines that don't report a real model now show an honest "model not reported by this device" instead of placeholder junk like "System Product Name."
- Fixed the custody form template's title banner ("CUSTODY/TRACKING FORM...") sometimes looking cramped when previewed inside OneDrive's browser view (the downloaded/emailed file itself was always fine - this was purely a web-preview rendering quirk, now resolved with a small header row height adjustment).

**v2.8** (2026-07-22)
- Added an "Is this a computer or laptop?" popup that appears first. Choosing "Yes" works exactly as before (auto-detects everything). Choosing "No" skips hardware auto-detection and lets you type in the item's category and description by hand - for things like projectors, clickers, or screens that you're custodying from your own computer rather than running the tool on directly.
- Smarter model name detection for Lenovo devices - now shows a clean, human-readable model name (e.g. "Model: Lenovo ThinkPad E14 Gen 7 (21SX001TGR)") instead of the raw internal SKU code, while still keeping the exact model number for asset records.
- Replaced the typed console prompts with a proper popup window. Location is now a dropdown limited to the four valid options (Mall of Qatar, Bani Hajer, Boys School, Girls School) so it can't be mistyped; the form won't close until all required fields are filled in.

**v2.7** (2026-07-22)
- Emailing now works no matter which staff member's account signs in - uploads always target jcarlos's OneDrive folder specifically (a hardcoded Drive ID), instead of whoever-signed-in's own OneDrive. Requires the CustodyFormsToEmail folder to be shared org-wide with edit access (see Email Automation above).

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
