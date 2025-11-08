# Mace Style Validator - User Guide

## Introduction

The Mace Style Validator automatically checks your Word documents against the Mace Control Centre Writing Style Guide. It identifies issues and automatically fixes many common errors, ensuring consistent, professional documentation.

---

## Getting Started

### What Gets Validated?

The validator checks for:

✅ **British English Spelling**
- finalized → finalised
- color → colour
- center → centre
- organize → organise
- And 20+ more...

✅ **Grammar Rules**
- Contractions expanded (can't → cannot, don't → do not)
- Proper apostrophe usage
- No apostrophes in plurals

✅ **Punctuation & Symbols**
- Ampersands replaced (& → and)
- Percent spelled out (% → percent)
- Number formatting with commas (1000 → 1,000)

✅ **Font Consistency**
- All text standardized to Arial
- Consistent heading fonts

---

## How to Validate a Document

### Method 1: Upload a New Document

1. **Navigate to your SharePoint Document Library**
   - Go to: `https://[yoursite].sharepoint.com/sites/StyleValidation`
   - Open the **Documents** library

2. **Upload your Word document**
   - Click **Upload** → **Files**
   - Select your `.docx` file
   - Click **Open**

3. **Automatic validation starts**
   - Status changes to **"Validating..."**
   - Typically takes 5-15 seconds

4. **Check the results**
   - Status updates to **"Passed"** (green) or **"Failed"** (red)
   - Click the **Validation Report** link to see details

### Method 2: Modify an Existing Document

1. **Check out the document** (optional but recommended)
   - Right-click → **Check Out**

2. **Edit in Word Desktop or Word Online**
   - Make your changes
   - Save the document

3. **Validation triggers automatically**
   - Status changes to **"Validating..."**
   - Results appear within seconds

---

## Understanding Validation Results

### Document Library Columns

Your document will show:

| Column | Description | Example |
|--------|-------------|---------|
| **Name** | Document filename | `Project_Report.docx` |
| **ValidationStatus** | Current status | 🟢 Passed / 🔴 Failed |
| **ValidationResultLink** | Link to detailed results | 📋 View Validation Result |
| **Modified** | Last validation date | 11/08/2025 8:42 PM |

### Status Meanings

🟢 **Passed**
- All issues were automatically fixed
- Document meets style guide requirements
- Safe to distribute

🟡 **Validating...**
- Validation in progress
- Wait a few seconds

🔴 **Failed**
- Some issues couldn't be fixed automatically
- Review the HTML report for details
- Manual corrections needed

---

## Reading the HTML Report

### Accessing the Report

**Option 1: From Document Library**
1. Find your document in the library
2. Click the **Validation Report** link in the row

**Option 2: From Validation Results List**
1. Go to **Validation Results** list
2. Click on your validation entry
3. Click the **Report Link**

### Report Structure

#### 1. Header Section
```
📋 Style Validation Report
[PASSED / FAILED Badge]

Document: Project_Report.docx
Validated: 08 November 2025 at 20:42:15 UTC
```

#### 2. Summary Dashboard
```
┌─────────────┬─────────────┬────────────────┐
│ Issues Found│ Auto-Fixed  │Remaining Issues│
│     156     │     156     │        0       │
└─────────────┴─────────────┴────────────────┘
```

#### 3. Fixes Applied (Green Section)
Lists all automatic corrections:

```
✅ Fixes Applied (156)

✓ Fixed 145 text runs to Arial
✓ Applied 8 style corrections (British English, contractions, symbols, etc.)
✓ Replaced 'finalized' with 'finalised' (3 instances)
```

#### 4. Issues Detected (Red Section)
Lists all problems found (including those fixed):

```
⚠️ Issues Detected (156)

⚠ Found 145 text runs with incorrect font
⚠ Found 8 style violations
```

### Report Color Coding

| Color | Meaning |
|-------|---------|
| 🟢 Green | Successfully fixed |
| 🔴 Red/Pink | Issue detected (may or may not be fixed) |
| 🟣 Purple | Status badge (Passed/Failed) |

---

## Validation Results List

### Viewing All Validations

1. Navigate to **Validation Results** list
   - Click **Validation Results** in site navigation
   - Or: `https://[yoursite].sharepoint.com/sites/StyleValidation/Lists/Validation Results`

2. **View validation history**
   - See all past validations
   - Filter by document name, date, or status
   - Track validation trends

### List Columns

| Column | Description |
|--------|-------------|
| **Title** | "Validation: {filename}" |
| **FileName** | Original document name |
| **ValidationDate** | When validated |
| **Status** | Passed / Failed |
| **IssuesFound** | Count of problems detected |
| **IssuesFixed** | Count of auto-fixes |
| **ReportLink** | Link to detailed HTML report |

### Filtering & Sorting

**Common filters:**
- Show only failed validations: `Status equals Failed`
- Recent validations: Sort by `ValidationDate` descending
- Specific document: Filter by `FileName`

---

## What Happens to Your Document?

### Automatic Fixes Applied

When issues are detected and `AutoFix = Yes`:

1. **Original document is backed up**
   - SharePoint versioning preserves history
   - You can always restore previous versions

2. **Fixes are applied**
   - Text corrections (spelling, grammar)
   - Font standardization
   - Symbol replacements

3. **Document is saved**
   - Overwrites the original file
   - Version history maintained

### Manual Review Needed

When `AutoFix = No` or issues can't be corrected:

1. **Document is not modified**
   - Original remains untouched

2. **Issues are reported**
   - Listed in HTML report
   - Status shows "Failed"

3. **You must fix manually**
   - Open document in Word
   - Review HTML report for specific issues
   - Make corrections
   - Save to re-trigger validation

---

## Common Scenarios

### Scenario 1: Document Passes Validation

✅ **What happened:**
- All issues auto-fixed
- Document meets style guide
- Ready for distribution

**Next steps:**
- No action needed
- Review HTML report to see what was fixed (optional)
- Download/share the document

---

### Scenario 2: Document Fails Validation

❌ **What happened:**
- Some issues couldn't be fixed automatically
- Manual review required

**Next steps:**
1. Open the HTML report
2. Review "Issues Detected" section
3. Open document in Word
4. Make manual corrections
5. Save → validation re-runs automatically

**Common manual fixes:**
- Date format corrections (must be "DD MONTH YYYY")
- Numbers below 10 spelled out in text
- Context-dependent grammar (licence vs. license)

---

### Scenario 3: Validation Takes Too Long

⏳ **If stuck on "Validating..." for >2 minutes:**

1. **Check file size**
   - Large documents (>50 pages) take longer
   - Consider splitting very large files

2. **Refresh the page**
   - Press F5 or click refresh
   - Status should update

3. **Check for errors**
   - Admin: Review Azure Function logs
   - Look for API timeouts or errors

---

### Scenario 4: Unwanted Changes Made

↩️ **To restore previous version:**

1. Click **...** (three dots) next to document
2. Select **Version History**
3. Find the version before validation
4. Click **...** → **Restore**

**Note:** Validation creates a new version each time it modifies the document.

---

## Best Practices

### ✅ Do's

✅ **Use descriptive filenames**
- Good: `Project_Alpha_Requirements.docx`
- Bad: `doc1.docx`

✅ **Review HTML reports**
- Understand what's being changed
- Learn style guide rules

✅ **Check version history**
- Verify changes before distributing
- Restore if needed

✅ **Fix root causes**
- If you consistently get same errors, update your templates
- Train team on common issues

✅ **Keep documents under 100 pages**
- Faster validation
- Better performance

### ❌ Don'ts

❌ **Don't disable SharePoint versioning**
- Needed for backup/restore

❌ **Don't upload unsupported formats**
- Only `.docx` and `.doc` supported
- Not `.pdf`, `.txt`, or other formats

❌ **Don't edit during validation**
- Wait for "Validating..." to complete
- Concurrent edits may cause conflicts

❌ **Don't ignore "Failed" status**
- Manual review required
- Document may not meet standards

---

## Troubleshooting

### Problem: Validation not triggering

**Possible causes:**
1. Wrong file format (not .docx/.doc)
2. Power Automate flow disabled
3. Permissions issue

**Solutions:**
- Verify file extension is `.docx`
- Contact admin to check flow status
- Ensure you have edit permissions

---

### Problem: Status stuck on "Validating..."

**Possible causes:**
1. Azure Function timeout
2. Large document
3. API error

**Solutions:**
- Wait 2 minutes, then refresh
- Check document size (<50 pages ideal)
- Contact admin to check logs

---

### Problem: Validation Report link broken

**Possible causes:**
1. Report upload failed
2. Permissions issue
3. SharePoint storage full

**Solutions:**
- Re-save document to trigger re-validation
- Contact admin to check Azure Function logs
- Verify SharePoint storage quota

---

### Problem: Unwanted changes made

**Possible causes:**
1. Style rule too aggressive
2. Misunderstanding of rule
3. Bug in validation logic

**Solutions:**
- Restore from version history
- Contact admin to review specific rule
- Report issue with examples

---

## Managing Style Rules

### Viewing Current Rules

1. Navigate to **Style Rules** list
2. See all active validation rules
3. Sort by Priority to see execution order

**Note:** Only administrators can modify rules.

### Requesting Rule Changes

If you believe a rule should be changed:

1. **Gather examples**
   - Document name
   - Specific text that was incorrectly changed
   - Expected vs. actual result

2. **Contact administrator**
   - Provide examples
   - Explain business justification

3. **Admin reviews and updates**
   - Rule modified in Style Rules list
   - Change takes effect immediately
   - All future validations use new rule

---

## Tips for Success

### 1. Learn the Style Guide
- Review common corrections in your reports
- Keep a cheat sheet of British spellings
- Understand contraction rules

### 2. Use Templates
- Create Word templates with correct fonts
- Pre-set styles (Heading 1, Normal, etc.)
- Start with compliant formatting

### 3. Validate Early and Often
- Don't wait until document is complete
- Validate as you draft
- Catch issues early

### 4. Review Before Distributing
- Always check the HTML report
- Verify corrections make sense
- Spot-check important sections

### 5. Provide Feedback
- Report false positives
- Suggest new rules
- Help improve the system

---

## Getting Help

### Self-Service Resources
1. **HTML Report** - Detailed issue descriptions
2. **Version History** - Restore previous versions
3. **Validation Results List** - Historical data

### Contact Support
- **Email:** [your-it-support@email.com]
- **Teams:** [IT Support Channel]
- **Phone:** [Support Number]

### Report Issues
Include:
- Document name
- Validation date/time
- Expected vs. actual behavior
- Screenshots if helpful

---

## Appendix: Supported Rules

### British English Spellings
| ❌ American | ✅ British |
|------------|-----------|
| finalized | finalised |
| color | colour |
| center | centre |
| organize | organise |
| analyze | analyse |
| defense | defence |
| fiber | fibre |
| gray | grey |
| harbor | harbour |
| labor | labour |
| meter | metre |
| + 15 more... | |

### Contractions Expanded
| ❌ Contraction | ✅ Expanded |
|---------------|-------------|
| can't | cannot |
| don't | do not |
| isn't | is not |
| won't | will not |
| couldn't | could not |
| didn't | did not |
| doesn't | does not |
| hasn't | has not |
| shouldn't | should not |
| wouldn't | would not |

### Symbols Replaced
| ❌ Symbol | ✅ Text |
|----------|---------|
| & | and |
| % | percent |

### Font Rules
- **All text:** Arial
- **All headings:** Arial

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| v4.2 | Nov 2025 | Current version with enhanced HTML reports |
| v3.3 | Nov 2025 | Added AI validation and Validation Results list |
| v2.0 | Oct 2025 | Added British English rules |
| v1.0 | Sep 2025 | Initial release with font validation |
