---
name: mail-drafter
description: >
  Collects mail subject, body, and recipients from the user, optionally attaches a logo (placed after "Regards" inline) and a file,
  then opens 3 preview drafts in Outlook Classic for review, asks if changes are needed or to auto-send all remaining emails.
---

# Mail Drafter Skill

## Purpose
Compose a professional email, open the first 3 drafts in Outlook Classic for visual review, then auto-send all remaining emails after user approval — one email per recipient, never combined.

---

## Step-by-Step Flow

### Step 1 — Collect Mail Details
Ask the user to provide the following:

- **Subject**: The email subject line.
- **Body**: Full email body text. The user can provide it as plain text or HTML. If plain text, preserve paragraph breaks.

### Step 2 — Ask About Attachments and Logo
Ask the user two questions:

1. **Logo**: "Do you want to add a logo/image after the Regards section? (yes/no)"
   - If **yes**: Ask for the full file path to the logo image (PNG, JPG, GIF supported).
   - Validate the file exists using `Test-Path`. If not found, ask again.

2. **File Attachment**: "Do you want to attach a file to the email? (yes/no)"
   - If **yes**: Ask for the full file path of the file to attach.
   - Validate the file exists using `Test-Path`. If not found, ask again.

### Step 3 — CC Recipient Confirmation
Prompt the user with the following:
"Would you like to add any CC recipients?"

**Options:**
- **Yes, Add CC**: Prompt for one or more CC email addresses. Validate format. Allow multiple IDs (comma/semicolon separated). Store these CC recipients as part of the email metadata.
- **No, Continue**: Skip this step.

### Step 4 — Collect "To" Recipients
Ask the user for the primary recipients:
- **Recipients**: One or more email addresses. Accept comma-separated, semicolon-separated, or line-by-line input.

Display a confirmation summary before proceeding:
```
Subject   : <subject>
Body      : <first 100 chars>...
CC        : <list each one> (or "None")
Recipients: <list each one, numbered>
Total     : <count of "To" recipients>
```

### Step 5 — Build the HTML Email Body
Construct a clean HTML body:

1. Convert the user's plain text body to HTML (replace `\n\n` with `</p><p>`, `\n` with `<br>`).
2. **Regards spacing fix**: Detect the sign-off block (case-insensitive match for "regards", "best regards", "warm regards", "thanks and regards", etc.).
   - Strip any extra blank lines or `<br>` tags between the last body sentence and the sign-off line — the regards block must sit flush against the line above it with exactly one `<br>` separation.
   - If logo is provided: place the `<img>` tag on the very next line after the sign-off, with NO blank line between them.
   - If NOT found: append the sign-off and logo at the end with one `<br>` gap from the body.
3. Logo sizing: use `width="150"` (150px wide, auto height) — professional and not oversized.
4. The logo must be embedded as a **CID inline attachment** (not a linked URL) so it renders inside the email body in Outlook.

Correct HTML structure (no extra spacing around Regards):
```html
<html>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;">
<p>{body_paragraphs}</p>
Best regards,<br>
{sender_name}<br>
<img src="cid:logo_inline" width="150" alt="Logo" style="display:block;margin-top:6px;">
</body>
</html>
```
If no logo is provided, omit the `<img>` tag entirely. Never insert a `<br>` or blank line between the sign-off text and the logo image.

### Step 6 — Preview 3 Drafts + Review Gate

This is a two-phase execution:

#### Phase A — Open First 3 Drafts for Review
1. Take the first 3 recipients from the list (or all if fewer than 3).
2. Open each as a draft using `.Display()` in Outlook so the user can visually inspect subject, body, logo placement, and attachment.
3. After opening the 3 drafts, pause and ask the user in the terminal:

```
3 preview drafts are now open in Outlook Classic.
Please review them — check the body, logo position, and attachment.

What would you like to do?
  [1] Everything looks good — auto-send ALL remaining emails
  [2] I need to make changes — let me update the details first
  [3] Cancel — discard everything
```

#### Phase B — Based on User Choice

**Choice 1 — Send All:**
- Auto-send the 3 preview drafts that are still open (call `.Send()` on their MailItem objects, or recreate and send if already closed).
- Create and auto-send all remaining recipients (recipient 4 onwards) using `.Send()`.
- Report progress in the terminal as emails go out (e.g., "Sent 1/100 → email@example.com").
- Add a 1-second delay between sends to avoid Outlook throttling.
- Final report: total sent, any failures.

**Choice 2 — Make Changes:**
- Ask the user what they want to change: subject / body / logo path / attachment path / recipients / CC recipients.
- Apply the changes.
- Close ALL currently open draft windows before proceeding (release their COM objects / call `.Close(olDiscard)` on each open MailItem).
- Re-run from Step 5 (rebuild HTML body) and open 3 brand-new fresh draft windows so the user sees clean tabs with only the updated content.
- Repeat the review gate.

**Choice 3 — Cancel:**
- Close any open draft windows (release COM objects).
- Confirm cancellation: "All drafts discarded. No emails were sent."

#### PowerShell Script Template

```powershell
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

$outlook   = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$recipients   = @(
    "recipient1@example.com",
    "recipient2@example.com"
    # ... more recipients
)
$ccRecipients = "cc1@example.com; cc2@example.com" # Semicolon separated string, or $null

$subject  = "Your Subject Here"
$htmlBody = @"
<html><body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;">
<p>Body content here.</p>
Best regards,<br>
[Sender Name]<br>
<img src="cid:logo_inline" width="150" alt="Logo" style="display:block;margin-top:6px;">
</body></html>
"@

$logoPath = "C:\path\to\logo.png"   # $null if no logo
$filePath = "C:\path\to\file.pdf"   # $null if no attachment

function New-MailItem($recipient, $display) {
    $mail = $outlook.CreateItem(0)
    $mail.Subject = $subject
    $mail.To      = $recipient
    
    if ($ccRecipients) {
        $mail.CC = $ccRecipients
    }

    if ($logoPath -and (Test-Path $logoPath)) {
        $att = $mail.Attachments.Add($logoPath, 1)
        $att.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            "logo_inline"
        )
        $mail.HTMLBody = $htmlBody   # Set AFTER inline attachment
    } else {
        $cleanBody = $htmlBody -replace '<img[^>]*logo_inline[^>]*>', ''
        $mail.HTMLBody = $cleanBody
    }

    if ($filePath -and (Test-Path $filePath)) {
        $mail.Attachments.Add($filePath, 1) | Out-Null
    }

    if ($display) {
        $mail.Display()   # Preview drafts — do NOT send
    } else {
        $mail.Send()      # Auto-send remaining
        Start-Sleep -Seconds 1
    }
    return $mail
}

# Phase A: Open first 3 as preview drafts
$previewRecipients  = $recipients[0..([Math]::Min(2, $recipients.Count - 1))]
$remainingRecipients = if ($recipients.Count -gt 3) { $recipients[3..($recipients.Count - 1)] } else { @() }

foreach ($r in $previewRecipients) {
    New-MailItem -recipient $r -display $true | Out-Null
}

Write-Host "3 preview drafts opened. Review in Outlook, then enter your choice:"
Write-Host "[1] Send all  [2] Make changes  [3] Cancel"
$choice = Read-Host "Choice"

if ($choice -eq "1") {
    # Recreate and send preview recipients (drafts may have been edited/closed)
    foreach ($r in $previewRecipients) {
        New-MailItem -recipient $r -display $false | Out-Null
        Write-Host "Sent -> $r"
    }
    $i = $previewRecipients.Count + 1
    foreach ($r in $remainingRecipients) {
        New-MailItem -recipient $r -display $false | Out-Null
        Write-Host "Sent $i/$($recipients.Count) -> $r"
        $i++
    }
    Write-Host "All $($recipients.Count) emails sent successfully."
} elseif ($choice -eq "3") {
    Write-Host "Cancelled. No emails were sent."
} else {
    Write-Host "Please update your details and re-run the skill."
}
```

**Important implementation notes:**
- Always set `HTMLBody` **after** adding the inline logo attachment — Outlook resets the body when set before.
- Phase A always uses `.Display()`, Phase B send-all always uses `.Send()`.
- The 1-second delay between sends prevents Outlook throttling on large batches.
- The CID string `"logo_inline"` in `<img src="cid:logo_inline">` must exactly match the value set via `SetProperty` (`0x3712001F`).
- For Choice 1, recreate the 3 preview mails as new MailItems and send — do not try to `.Send()` already-displayed drafts as their state may be changed by the user.

### Step 7 — Confirm and Execute
Before running any script, show the user a final summary:

```
Ready to run:
  Recipients  : <total count> individual emails
  CC          : <list each one> (or "None")
  Subject     : <subject>
  Logo        : <filename> (inline after Regards) OR None
  Attachment  : <filename> OR None
  Preview     : First 3 drafts will open in Outlook for your review
  After review: You choose to send all or make changes

Proceed? (yes/no)
```

If confirmed, write the PowerShell script to `$env:TEMP\mail_drafter_<timestamp>.ps1` and run it via:
```powershell
powershell.exe -ExecutionPolicy Bypass -File "$env:TEMP\mail_drafter_<timestamp>.ps1"
```

---

## Error Handling

| Scenario | Action |
|---|---|
| Outlook not installed | Tell the user Outlook Classic is required and stop. |
| Logo file not found | Warn and ask for correct path; offer to skip logo. |
| Attachment file not found | Warn and ask for correct path; offer to skip attachment. |
| Zero recipients provided | Ask again — do not proceed with empty recipient list. |
| Send failure on a recipient | Log the failure, continue with remaining, report at end. |
| Script execution policy error | Suggest: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` |

---

## Constraints
- NEVER auto-send without the review gate (Phase A preview + user Choice 1).
- NEVER combine multiple recipients into one `To` field — one email per person.
- NEVER skip the confirmation summary in Step 7.
- NEVER insert extra blank lines or `<br>` tags between the sign-off text and the logo image.
- Logo must be inline (CID), not a file path reference or URL.
- NEVER polish, modify, rewrite, or alter the user's pasted content in any way — use it exactly as provided, character for character, preserving all original formatting, wording, punctuation, and paragraph breaks.
- When the user requests changes after reviewing the preview drafts: close ALL open draft windows first, then open fresh new draft windows with the updated content applied — never reuse or edit the existing draft windows.
