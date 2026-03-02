# Personal Injury Paralegal - Claude Plugin

A Claude plugin that automates the full demand letter workflow for personal injury law firms. One prompt drafts the letter, sends it, calls the adjuster, and documents everything, all from inside your existing systems.

## How It Works

You give Claude a single instruction, "Draft a demand letter for matter 1234-James", and the AI agent:

1. Pulls all case data from Clio Manage (matter details, medical records, damages, notes, police report)
2. Fills your firm's Word template with tracked changes
3. Presents the draft for your review and edits in Word
4. Generates a clean PDF once approved
5. Uploads final documents to Clio Manage
6. Emails the demand to the adjuster with the PDF attached
7. Calls the adjuster via AI voice agent to confirm receipt
8. Logs a matter note and creates a follow-up task in Clio

Nothing goes out until you review and approve it. Every edit appears as tracked changes in Word.

## Project Structure

```
├── .claude-plugin/
│   └── plugin.json                     # Plugin metadata (name, version, author)
├── .mcp.json                           # MCP server connection to Make.com
├── skills/
│   └── draft-demand-letter/
│       ├── SKILL.md                    # ⭐ THE MAIN FILE: the step-by-step workflow
│       └── templates/
│           └── demand-letter-template.docx  # Word template with placeholders
├── scripts/                            # Reusable Python helper scripts
│   ├── docx-template-editor.py         # Word template editing with tracked changes
│   ├── docx-accept-tracked-changes.py  # Strips tracked change markup for clean PDF
│   ├── docx-convert-to-pdf.py          # Word-to-PDF conversion (crash-tolerant)
│   ├── clio-manage-download-document.py # Downloads documents from Clio via Make.com webhook
│   └── clio-manage-upload-document.py   # Uploads documents to Clio via Make.com webhook
└── references/
    └── docx-template-editing-reference.md  # Guide for editing .docx templates with tracked changes
```

## Where to Start

**The most important file is [`skills/draft-demand-letter/SKILL.md`](skills/draft-demand-letter/SKILL.md).** This is the 15-step workflow that tells Claude exactly what to do. Read this first to understand the full process.

If you want to adapt this for your firm:

1. **Read the Skill file** to understand the workflow
2. **Replace the Word template** with your firm's demand letter template, adding `{PLACEHOLDER}` markers and `{{AI SECTION}}` blocks
3. **Build Make.com scenarios** for your CMS (you'll need scenarios for reading case data, uploading/downloading documents, sending email, and optionally phone calls)
4. **Connect Make.com to the plugin** via MCP (add the MCP server URL to `.mcp.json`)
5. **Update the Skill file** to match how your firm organizes data in your CMS
6. **Run the workflow** and iterate: let Claude improve the Skill file based on what it encounters

## Adapting for Your Firm

This plugin is built for Clio Manage, but the architecture works with any CMS. To adapt:

- **Different CMS (MyCase, Filevine, SmartAdvocate, etc.):** Replace the Clio Make.com scenarios with equivalent ones for your CMS. Make.com has pre-built connectors for most legal software. Update the Skill file's step descriptions accordingly.
- **Different template:** Replace `demand-letter-template.docx` with your template. Add `{PLACEHOLDER}` markers for data fields and `{{AI SECTION START/END}}` markers for narrative sections.
- **Different email provider:** Replace the Gmail scenario with one for your email system.
- **No phone calls:** Remove Steps 13a-13c from the Skill file. Everything else works without it.
- **Local files instead of API:** If you use Clio Drive, Google Drive, or OneDrive with desktop sync, you can skip the upload/download scripts entirely and point the agent to the synced folder on your machine.
