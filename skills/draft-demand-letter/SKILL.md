---
name: draft-demand-letter
description: "Specify the Matter Display Number (e.g. '1234-James'). Draft a personal injury demand letter. Use when to draft, create, or send a demand letter for a Clio Manage matter."
---

# Draft Demand Letter

You are a paralegal assistant for a personal injury law firm. Draft a demand letter for matter **$ARGUMENTS** by following the steps below exactly in order.

## Critical Rules

- NEVER fabricate case facts, medical records, damage amounts, or any other data. Every data point must come from Clio Manage records or the police report.
- All monetary amounts must use the exact figures from Clio Manage damages records.
- Document download and upload MUST use the provided Python scripts (deterministic binary handling). Do NOT attempt to handle binary file data directly.
- If any required data is missing, STOP and ask the user before proceeding. Do not guess or fill in defaults.
- The .docx file on disk is ALWAYS the source of truth. Before any edit or conversion, re-read the file from disk. The user may modify the file externally (accepting tracked changes, making manual edits) at any time between messages. Never rely on cached or in-memory state from previous steps.

## Step 1: Retrieve Matter

Call the `make-pi-paralegal` MCP tool `clio_manage_get_matter_fields` with the matter display number "$ARGUMENTS" to get the matter record.

Expected input: `{"matter_display_number": "$ARGUMENTS"}`

Expected output: All matter fields including `matter_id`, client name, client address, opposing party name, insurance company, adjuster name, adjuster email, adjuster phone number, and other case details.

Extract and store the `matter_id` for all subsequent calls. If the matter is not found, inform the user and stop.

## Step 2: Gather Additional Matter Information

Call the following `make-pi-paralegal` MCP tools using the `matter_id`:

1. **`clio_manage_get_matter_medical_records`** - Input: `{"matter_id": "<ID>"}` - Returns: treatment history (provider names, dates, treatment types, costs)
2. **`clio_manage_get_matter_damages`** - Input: `{"matter_id": "<ID>"}` - Returns: injuries sustained and damage amounts (medical expenses, lost wages, pain & suffering, etc.)
3. **`clio_manage_get_matter_notes`** - Input: `{"matter_id": "<ID>"}` - Returns: staff notes (additional accident details, non-economic damages like loss of enjoyment of life, other case remarks)

Store all results for template field mapping in Step 5.

## Step 3: Get Police Report

Call the `make-pi-paralegal` MCP tool **`clio_manage_get_matter_documents`** - Input: `{"matter_id": "<ID>"}` - Returns: document list with IDs, names, and types.

From the returned list, identify the police report (look for document names containing "police", "accident report", "crash report", "traffic report", or similar).

If found, download it:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/clio-manage-download-document.py" \
  --document-id <DOCUMENT_ID> \
  --output-path ./police-report.pdf
```

Then read the downloaded PDF to extract: accident date, location, description of the accident, and liability details.

If no police report is found, ask the user whether to proceed without it or provide the document ID manually.

## Step 4: Read Template and Discover Required Fields

Read the demand letter template to discover all placeholder fields:

`${CLAUDE_PLUGIN_ROOT}/skills/draft-demand-letter/templates/demand-letter-template.docx`

Identify all `{XXX}` style placeholders (e.g., `{CLIENT_NAME}`, `{DATE_OF_INCIDENT}`, `{TOTAL_MEDICAL_EXPENSES}`, etc.). Some placeholders may be dynamic sections, that will require your reasoning and a write-up later on (like the liability section). Identify what information pieces may be required for those.

Build a complete list of all fields & information pieces required to draft from the template.

## Step 5: Map Data to Placeholders

For each placeholder field & information piece discovered in Step 4, map it to a value from the data gathered in Steps 1-3:

- Matter fields from Step 1 (client info, opposing party, adjuster contact, etc.)
- Medical records from Step 2 (treatment history, provider details)
- Damages from Step 2 (injury descriptions, amounts)
- Notes from Step 2 (additional case details, non-economic damages)
- Police report facts from Step 3 (accident date, location, description, liability)

Create a complete mapping of every placeholder to its value.

## Step 6: Check for Missing Data

Compare the required placeholder fields and information pieces against your mapping. If ANY fields or information pieces are missing values:
Present the user with the missing data point & ask the user to provide a value for it (with the native Ask User functionality).

Wait for the user's response before continuing.

Repeat this process for each missing data point.

Do NOT proceed with any missing data points.

## Step 7: Draft Demand Letter

BEFORE writing any code to edit the template, read the DOCX template editing reference guide:

`${CLAUDE_PLUGIN_ROOT}/references/docx-template-editing-reference.md`

This guide covers the exact step-by-step process for editing .docx templates with tracked changes, including all known pitfalls and how to avoid them.

Copy the template into the current working directory:

```bash
cp "${CLAUDE_PLUGIN_ROOT}/skills/draft-demand-letter/templates/demand-letter-template.docx" ./demand-letter-draft.docx
```

Then write a Python script to populate the template. Import the reusable editor:

```python
import sys, importlib
sys.path.insert(0, "${CLAUDE_PLUGIN_ROOT}/scripts")
docx_editor = importlib.import_module("docx-template-editor")
edit_template = docx_editor.edit_template
```

Use `edit_template()` to replace all `{PLACEHOLDER}` fields and `{{AI SECTION}}` blocks with the actual values from your mapping. The editor handles tracked changes, highlight removal, paragraph spacing, and all Word XML edge cases automatically.

Follow the reference guide for the full workflow: unpack → edit → repack → verify.

## Step 8: User Review and Revisions

Present the .docx draft (with tracked changes) for the user's review. **Only share the .docx file — do NOT convert to PDF or share a PDF at this stage.** The user may open the .docx in Word, accept/reject tracked changes, and make their own edits.

**IMPORTANT:** The user edits the file in the outputs/workspace folder (the version you shared via the `computer://` link). Before ANY subsequent read, unpack, or conversion of the .docx, you MUST first copy the file FROM the outputs folder BACK to your working directory:

```bash
cp <outputs_folder>/demand-letter-draft.docx ./demand-letter-draft.docx
```

This applies to Step 8 revision rounds AND Step 9 PDF generation. The working directory copy is stale after the user edits the outputs version.

If the user requests any changes:

1. **Copy the user's version back from the outputs folder**, then re-read the .docx file from disk. The user may have accepted tracked changes, made manual edits, or otherwise modified the file outside of this session. Never assume the unpacked XML from Step 7 is still current.
2. Re-unpack the .docx: `python3 .../unpack.py ./demand-letter-draft.docx ./unpacked/`
3. Re-analyze the unpacked XML to understand the current document state (spacing, rPr, content).
4. Apply the requested change to the freshly unpacked XML.
5. Repack: `python3 .../pack.py ./unpacked/ ./demand-letter-draft.docx --original <template> --validate false`
6. Present the updated .docx to the user.
7. Repeat until the user confirms the draft is approved.

**CRITICAL:** Never edit stale XML. Every revision round starts with a fresh unpack of the current .docx on disk.

## Step 9: Generate Final Clean PDF

Once the user confirms the draft is approved and no further changes are needed:

1. **Copy the user's version back from the outputs folder:**
   ```bash
   cp <outputs_folder>/demand-letter-draft.docx ./demand-letter-draft.docx
   ```
2. **Accept all remaining tracked changes** programmatically before converting to PDF. The user may not have accepted all changes themselves, and LibreOffice's PDF export renders tracked changes as visible markup (strikethrough, underlines, colored text). The final PDF must be clean.
   ```bash
   python3 "${CLAUDE_PLUGIN_ROOT}/scripts/docx-accept-tracked-changes.py" ./demand-letter-draft.docx ./demand-letter-clean.docx
   ```
3. Convert the clean .docx to PDF using the crash-tolerant wrapper:
   ```bash
   python3 "${CLAUDE_PLUGIN_ROOT}/scripts/docx-convert-to-pdf.py" ./demand-letter-clean.docx ./demand-letter-final.pdf
   ```

This PDF is the **final clean version** that will be emailed to the adjuster as-is. No further conversions or modifications should be made after this point.

## Step 10: Final Review and Confirm Send Details

Present the final PDF to the user for a last review. This is the exact document that will be sent to the adjuster.

Then present the adjuster contact details:

- **At-Fault Claims Adjuster name**: (from matter fields in Step 1)
- **At-Fault Claims Adjuster email**: (from matter fields in Step 1)
- **At-Fault Claims Adjuster phone**: (from matter fields in Step 1)

Ask the user to confirm:
1. The PDF looks correct and is ready to send
2. The adjuster contact details are correct

Do NOT proceed to upload/email/call until both are confirmed.

## Step 11: Upload to Clio Manage

Upload both files to the matter in Clio Manage:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/clio-manage-upload-document.py" \
  --matter-id <MATTER_ID> \
  --file-path ./demand-letter-draft.docx \
  --document-name "Demand Letter Final.docx"
```

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/clio-manage-upload-document.py" \
  --matter-id <MATTER_ID> \
  --file-path ./demand-letter-final.pdf \
  --document-name "Demand Letter Final.pdf"
```

Capture the resulting Clio Manage document IDs.

## Step 12: Send Email to Adjuster

Call the `make-pi-paralegal` MCP tool `gmail_send_email`:

- **to**: Adjuster email confirmed in Step 10
- **subject**: `Demand Letter - [Client Name] v. [Opposing Party] - Matter [Display Number]`
- **body**: Draft a professional cover letter. The cover letter should briefly state that the attached demand letter is being submitted on behalf of the client. **IMPORTANT:** Do NOT include a signature block (attorney name, firm name, address, phone number) at the end of the email body. The Gmail integration automatically appends the sender's configured email signature. Including one manually will result in a duplicate signature. The email body should end with the closing line of the cover letter (e.g., "Should you have any questions, please do not hesitate to contact our office.") — no "Sincerely," block, no name, no address.
- **clio_manage_document_id**: The Clio Manage document ID of the uploaded PDF

Expected input:
```json
{
  "to": "<adjuster_email>",
  "subject": "<email_subject>",
  "body": "<cover_letter_text>",
  "clio_manage_document_id": "<pdf_document_id>"
}
```

## Step 13: Call Adjuster

### 13a. Start the call

Call the `make-pi-paralegal` MCP tool `retell_start_call`:

- **phone_number**: Adjuster phone confirmed in Step 10
- **instructions**: `You are calling the At-Fault Insurance Claims Adjuster [Adjuster Name] regarding the personal injury matter [Client Name] v. [Opposing Party], matter number [Display Number]. A demand letter was just sent via email to them. Please: (1) Confirm receipt of the demand letter email. After the confirmation: (2) Ask when the firm can expect a response.`

Expected input:
```json
{
  "phone_number": "<adjuster_phone>",
  "instructions": "<detailed_call_script>"
}
```

This returns a `call_id` immediately.

### 13b. Poll for completion

Loop up to 20 times:
1. Sleep 30 seconds: `sleep 30` via Bash
2. Call the `make-pi-paralegal` MCP tool `retell_get_call_status` with `{"call_id": "<call_id>"}`
3. If `call_status` is `"ended"`, extract the transcript and break
4. If `call_status` is `"ongoing"`, continue polling
5. If `call_status` is `"failed"`, log the error and break

Maximum wait: ~10 minutes (20 x 30s).

### 13c. Extract call details

Read the transcript to extract:
- Whether the adjuster confirmed receipt of the demand letter
- Any response timeline or follow-up date they committed to
- Any other relevant details mentioned during the call

Use these extracted details in the note and task steps below.

## Step 14: Create Matter Note

Call the `make-pi-paralegal` MCP tool `clio_manage_create_matter_note`:

- **matter_id**: The matter ID
- **subject**: Short subject line for the note (e.g., `Demand Letter Sent`)
- **content**: A summary including:
  - Demand letter drafted and finalized on [today's date]
  - Uploaded to Clio Manage (Document IDs: .docx = [ID], .pdf = [ID])
  - Emailed to [adjuster name] at [adjuster email]
  - Follow-up call made to [adjuster phone]
  - Call transcript summary: receipt confirmed / not confirmed, expected response timeline, any other relevant details from the transcript

Expected input:
```json
{
  "matter_id": "<matter_id>",
  "subject": "Demand Letter Sent",
  "content": "<note_text>"
}
```

## Step 15: Create Follow-Up Task

Call the `make-pi-paralegal` MCP tool `clio_manage_create_matter_task`:

- **matter_id**: The matter ID
- **name**: Short task title (e.g., `Follow up on demand letter`)
- **description**: `Follow up on demand: sent [today's date, YYYY-MM-DD] to [adjuster firm, adjuster name]`
- **due_date**: The date agreed with the adjuster during the call. If no specific date was agreed, default to 30 days from today.
- **assignee_user_id**: Use the responsible staff user ID field from Step 1.

Expected input:
```json
{
  "matter_id": "<matter_id>",
  "name": "Follow up on demand letter",
  "description": "<task_description>",
  "due_date": "<YYYY-MM-DD>",
  "assignee_user_id": "<user_id>"
}
```

## Summary

Present a final summary to the user:

- Demand letter: drafted, reviewed, finalized
- PDF and Word documents uploaded to Clio Manage
- Email sent to [adjuser firm, adjuster name, adjuster email]
- Call made to [adjuser firm, adjuster name, adjuster phone] - [status]
- Matter note created
- Follow-up task created - due [date], assigned to [assignee]
