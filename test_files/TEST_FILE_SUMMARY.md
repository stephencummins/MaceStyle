# Visio Structural Validation Test File

## Quick Start

### Option 1: Manual Setup (Comprehensive)

Follow the detailed instructions in: `VISIO_TEST_INSTRUCTIONS.txt`

This will create a file with **all possible violations** across:
- ✓ Page dimensions
- ✓ Shape sizes
- ✓ Shape positions
- ✓ Fonts
- ✓ Colors
- ✓ Text content

**Time required:** ~15-20 minutes

---

### Option 2: Quick Test (Minimal)

For a quick validation test:

1. **Open Microsoft Visio Desktop**
2. **Create new blank drawing**
3. **Change page size:**
   - Design → Size → More Page Sizes
   - Set to: 8.5" × 11.0" (Portrait)
4. **Add 3-5 shapes with text:**
   - Any rectangles or circles
   - Add text like: "This project was finalized and we can't proceed"
5. **Save as:** `test_structural_validation.vsdx`
6. **Upload to SharePoint**

**Time required:** ~2-3 minutes

---

## What Will Be Validated

When you upload the file, the validation will check:

### 1. Page Dimensions ✓
- **Rule:** All pages must be 11.0" × 8.5" (Letter Landscape)
- **Fix:** Resizes page to correct dimensions

### 2. Shape Sizes ✓
- **Rules:**
  - Title boxes: 3.0" × 1.0"
  - Icons: 0.5" × 0.5"
  - Process boxes: 2.0" × 1.5"
- **Fix:** Resizes shapes to standard dimensions

### 3. Shape Positions ✓
- **Rules:**
  - Top margin: Y ≤ 2.0"
  - Left margin: X ≥ 1.0"
  - Right margin: X ≤ 10.0"
  - Bottom margin: Y ≥ 1.0"
  - Logo exact position: (0.5", 7.5")
  - Footer exact position: (9.5", 0.5")
- **Fix:** Moves shapes to comply with margins/positions

### 4. Fonts ✓
- **Rule:** All text must be Arial
- **Fix:** Changes all fonts to Arial

### 5. Colors ✓
- **Rules:**
  - Shape fill: #003399 (brand blue)
  - Text color: #000000 (black)
- **Fix:** Updates colors to brand standards

### 6. Text Content (AI-Powered) ✓
- **Rules:** British English, proper grammar, formatting
- **Fixes:**
  - finalized → finalised
  - can't → cannot
  - M&S → M and S
  - 50% → 50 percent
  - 1000 → 1,000
  - And many more!

---

## Expected Validation Report

After uploading, you'll receive an HTML report showing:

```
📋 Style Validation Report
[PASSED Badge]

Document: test_structural_validation.vsdx
Validated: 2025-01-10 at 14:32:15 UTC

Summary
┌─────────────┬─────────────┬────────────────┐
│Issues Found │  Auto-Fixed │Remaining Issues│
│     68      │     68      │       0        │
└─────────────┴─────────────┴────────────────┘

✅ Fixes Applied (68)
  ✓ Resized 1 page to 11.0x8.5
  ✓ Resized 6 shapes to standard dimensions
  ✓ Repositioned 6 shapes to margins/exact positions
  ✓ Fixed 15 shapes to Arial
  ✓ Applied 25 text style corrections
  ✓ Updated 15 colors to brand standards

⚠️ Issues Detected (68)
  ⚠ Found 1 page with incorrect dimensions
  ⚠ Found 6 shapes with incorrect dimensions
  ⚠ Found 6 shapes with incorrect position
  ⚠ Found 15 shapes with incorrect font
  ⚠ Found 25 style violations
  ⚠ Found 15 shapes with incorrect colors
```

---

## Files in This Directory

| File | Purpose |
|------|---------|
| `VISIO_TEST_INSTRUCTIONS.txt` | Detailed step-by-step setup guide |
| `VISIO_LAYOUT_REFERENCE.txt` | Visual reference for shape placement |
| `SETUP_CHECKLIST.txt` | Quick checklist for creating test file |
| `TEST_FILE_SUMMARY.md` | This file - quick reference |

---

## Testing Tips

### To Test Specific Features:

**Page Dimensions:**
- Create file in portrait mode (8.5" × 11.0")

**Shape Sizes:**
- Create shapes larger or smaller than standards
- Vary by ±0.2" to trigger violations

**Margins:**
- Place shapes near page edges
- Use coordinate positions outside margin bounds

**Exact Positions:**
- Label a shape "Logo" or "Company Logo"
- Place it somewhere other than (0.5", 7.5")

**Text Content:**
- Use American spellings: "finalized", "color", "analyze"
- Use contractions: "can't", "won't", "don't"
- Use symbols: "M&S", "50%", "25%"
- Use unformatted numbers: "1000", "5000"

**Fonts:**
- Set text to Calibri, Times New Roman, or Verdana
- (Visio may auto-convert, but try anyway)

**Colors:**
- Fill shapes with red, green, yellow
- Set text to various colors

---

## Troubleshooting

### "File won't upload to SharePoint"
- Check file size (< 100 MB)
- Ensure .vsdx format (not .vsd)
- Check SharePoint permissions

### "Validation not running"
- Verify Power Automate flow is enabled
- Check Azure Function is deployed
- Review Application Insights logs

### "No fixes applied"
- Check AutoFix: Yes on rules
- Verify rules have correct DocumentType: Visio
- Check Priority values (lower = higher priority)

### "Tolerance not working"
- Add Tolerance column to Style Rules list
- Set values: 0.1 for sizes/margins, 0.05 for exact positions
- Re-run validation

---

## Sample Text Content

Copy-paste these into shapes to test AI validation:

**American Spelling:**
```
The project was finalized last week and requires authorization.
Our organization will analyze the color scheme at the center.
We need to optimize and organize the catalog.
```

**Contractions & Symbols:**
```
We can't proceed without M&S approval.
The report won't be available until they've confirmed 50% completion.
Johnson & Johnson agreed to a 25% increase for 5000 items.
```

**Numbers:**
```
Budget: 1000 initially, increased to 5000 for 10000 items.
Cost per unit: 2500 with discount of 500.
```

**Combined Violations:**
```
The finalized report can't be authorized until M&S confirms
50% completion for 5000 items at the organization center.
```

---

## Video Demo Script

If recording a demo:

1. **Show original file** with violations
2. **Upload to SharePoint** document library
3. **Wait for validation** (~15-30 seconds)
4. **Open validation report** - show issues and fixes
5. **Download corrected file**
6. **Compare before/after** - show changes
7. **Highlight key features:**
   - Page resized
   - Shapes repositioned
   - Text corrected
   - Fonts standardized
   - Colors updated

---

## Next Steps After Testing

Once validation works:

1. ✓ **Add more rules** for your specific needs
2. ✓ **Adjust tolerances** based on results
3. ✓ **Create templates** with pre-validated structure
4. ✓ **Document standards** for your team
5. ✓ **Train users** on validation system
6. ✓ **Monitor reports** for common violations
7. ✓ **Refine rules** based on feedback

---

## Support

For issues or questions:
- Review: `docs/visio-validation-guide.md`
- Check: Azure Function logs in Application Insights
- Review: Validation Results list in SharePoint
- Report: GitHub Issues

---

**Created:** 2025-01-10
**Version:** 1.0
**Status:** Ready for Testing
