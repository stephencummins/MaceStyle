# Visio Structural Validation Rules - SharePoint Examples

Complete rule configurations for implementing structural validation in your SharePoint Style Rules list.

## Prerequisites

Ensure your SharePoint "Style Rules" list has these columns:
- Title (Text)
- RuleType (Text)
- DocumentType (Choice: Word/Visio/Both)
- CheckValue (Text)
- ExpectedValue (Text)
- AutoFix (Yes/No)
- Priority (Number)
- Tolerance (Number) - **Add this column if missing**

## How to Add the Tolerance Column

If your Style Rules list doesn't have a Tolerance column:

1. Go to SharePoint → Style Rules list
2. Click **+ Add column** → Number
3. Name: `Tolerance`
4. Description: "Acceptable variance for dimension/position validation (in inches)"
5. Default value: `0.1`
6. Click **Save**

## Complete Structural Rule Set

### 1. Page Standardization Rules

#### Rule 1.1: Letter Size Landscape
```
Title: Visio - Page Size Letter Landscape
RuleType: PageDimensions
DocumentType: Visio
CheckValue: PageSize
ExpectedValue: 11.0x8.5
AutoFix: Yes
Priority: 80
UseAI: No
Description: Standardize all Visio pages to Letter size landscape (11" × 8.5")
```

#### Rule 1.2: A4 Size Landscape (Alternative)
```
Title: Visio - Page Size A4 Landscape
RuleType: PageDimensions
DocumentType: Visio
CheckValue: PageSize
ExpectedValue: 11.69x8.27
AutoFix: Yes
Priority: 80
UseAI: No
Description: Standardize all Visio pages to A4 landscape
```

### 2. Shape Size Rules

#### Rule 2.1: Title Box Dimensions
```
Title: Visio - Title Box 3"×1"
RuleType: Size
DocumentType: Visio
CheckValue: TitleBoxSize
ExpectedValue: 3.0x1.0
AutoFix: Yes
Priority: 90
Tolerance: 0.1
UseAI: No
Description: All title boxes must be 3 inches wide by 1 inch tall (±0.1")
```

#### Rule 2.2: Icon Standardization
```
Title: Visio - Icons 0.5" Square
RuleType: Size
DocumentType: Visio
CheckValue: IconSize
ExpectedValue: 0.5x0.5
AutoFix: Yes
Priority: 92
Tolerance: 0.05
UseAI: No
Description: All icon shapes must be 0.5" × 0.5" square
```

#### Rule 2.3: Process Box Size
```
Title: Visio - Process Box 2"×1.5"
RuleType: Size
DocumentType: Visio
CheckValue: ProcessBoxSize
ExpectedValue: 2.0x1.5
AutoFix: Yes
Priority: 91
Tolerance: 0.1
UseAI: No
Description: Standard process boxes must be 2" wide × 1.5" tall
```

#### Rule 2.4: Legend Box
```
Title: Visio - Legend 4"×2"
RuleType: Size
DocumentType: Visio
CheckValue: LegendSize
ExpectedValue: 4.0x2.0
AutoFix: Yes
Priority: 93
Tolerance: 0.2
UseAI: No
Description: Legend boxes should be 4" × 2"
```

### 3. Position Rules - Margins

#### Rule 3.1: Top Margin (Headers)
```
Title: Visio - Header Top Margin
RuleType: Position
DocumentType: Visio
CheckValue: TopMargin
ExpectedValue: 2.0
AutoFix: Yes
Priority: 85
Tolerance: 0.1
UseAI: No
Description: Header shapes must stay within top 2 inches of page
```

#### Rule 3.2: Left Margin
```
Title: Visio - Left Margin 1"
RuleType: Position
DocumentType: Visio
CheckValue: LeftMargin
ExpectedValue: 1.0
AutoFix: Yes
Priority: 86
Tolerance: 0.1
UseAI: No
Description: All shapes must start at or after 1" from left edge
```

#### Rule 3.3: Right Margin
```
Title: Visio - Right Margin 10"
RuleType: Position
DocumentType: Visio
CheckValue: RightMargin
ExpectedValue: 10.0
AutoFix: Yes
Priority: 87
Tolerance: 0.1
UseAI: No
Description: All shapes must end before 10" mark (1" right margin on 11" page)
```

#### Rule 3.4: Bottom Margin
```
Title: Visio - Bottom Margin 1"
RuleType: Position
DocumentType: Visio
CheckValue: BottomMargin
ExpectedValue: 1.0
AutoFix: Yes
Priority: 88
Tolerance: 0.1
UseAI: No
Description: All shapes must stay above 1" from bottom edge
```

### 4. Position Rules - Exact Placement

#### Rule 4.1: Company Logo Position
```
Title: Visio - Logo Top-Left
RuleType: Position
DocumentType: Visio
CheckValue: ExactPosition
ExpectedValue: 0.5,7.5
AutoFix: Yes
Priority: 75
Tolerance: 0.05
UseAI: No
Description: Company logo must be at exact position (0.5", 7.5") - top-left corner
```

#### Rule 4.2: Date/Version Footer
```
Title: Visio - Footer Bottom-Right
RuleType: Position
DocumentType: Visio
CheckValue: ExactPosition
ExpectedValue: 9.5,0.5
AutoFix: Yes
Priority: 76
Tolerance: 0.05
UseAI: No
Description: Date and version info at bottom-right (9.5", 0.5")
```

#### Rule 4.3: Diagram Title Centered
```
Title: Visio - Title Centered
RuleType: Position
DocumentType: Visio
CheckValue: ExactPosition
ExpectedValue: 5.5,7.5
AutoFix: Yes
Priority: 77
Tolerance: 0.1
UseAI: No
Description: Diagram title centered at top (5.5" is center of 11" page)
```

## How to Import These Rules

### Method 1: Manual Entry

1. Navigate to SharePoint → Style Validation → Style Rules
2. Click **+ New** for each rule
3. Copy values from the examples above
4. Fill in each column exactly as shown
5. Click **Save**

### Method 2: Quick Edit (Bulk Entry)

1. Navigate to SharePoint → Style Validation → Style Rules
2. Click **Quick edit** at the top
3. Copy-paste multiple rows at once
4. Click **Exit quick edit** when done

### Method 3: PowerShell Script

```powershell
# Connect to SharePoint
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/StyleValidation"

# Define rules
$rules = @(
    @{Title="Visio - Page Size Letter Landscape"; RuleType="PageDimensions"; DocumentType="Visio"; CheckValue="PageSize"; ExpectedValue="11.0x8.5"; AutoFix=$true; Priority=80},
    @{Title="Visio - Title Box 3x1"; RuleType="Size"; DocumentType="Visio"; CheckValue="TitleBoxSize"; ExpectedValue="3.0x1.0"; AutoFix=$true; Priority=90; Tolerance=0.1},
    @{Title="Visio - Icons 0.5 Square"; RuleType="Size"; DocumentType="Visio"; CheckValue="IconSize"; ExpectedValue="0.5x0.5"; AutoFix=$true; Priority=92; Tolerance=0.05},
    @{Title="Visio - Header Top Margin"; RuleType="Position"; DocumentType="Visio"; CheckValue="TopMargin"; ExpectedValue="2.0"; AutoFix=$true; Priority=85; Tolerance=0.1},
    @{Title="Visio - Logo Top-Left"; RuleType="Position"; DocumentType="Visio"; CheckValue="ExactPosition"; ExpectedValue="0.5,7.5"; AutoFix=$true; Priority=75; Tolerance=0.05}
)

# Add rules to list
foreach ($rule in $rules) {
    Add-PnPListItem -List "Style Rules" -Values $rule
    Write-Host "Added: $($rule.Title)"
}
```

## Testing Your Rules

### Create Test Visio File

1. Open Microsoft Visio
2. Create new blank drawing
3. Add shapes with various sizes and positions
4. Save as `test_structural_validation.vsdx`
5. Upload to SharePoint

### Expected Validation Results

When you upload the test file, the validator should:

1. **Resize page** to 11" × 8.5" (if different)
2. **Resize shapes** to match configured dimensions
3. **Reposition shapes** to comply with margins
4. **Move specific shapes** to exact positions
5. **Generate report** showing all fixes applied

### Review Validation Report

The HTML report will show:
- Found X pages with incorrect dimensions → Resized X pages
- Found X shapes with incorrect dimensions → Resized X shapes
- Found X shapes with incorrect position → Repositioned X shapes

## Common Configurations

### Configuration A: Strict Corporate Standards
```
- Page size: Letter Landscape (AutoFix: Yes)
- All title boxes: 3"×1" (Tolerance: 0.05")
- Logo position: Exact at 0.5", 7.5" (Tolerance: 0.05")
- Top margin: 2" (Tolerance: 0.1")
- Left/Right margins: 1" and 10" (Tolerance: 0.1")
```

### Configuration B: Flexible Layout
```
- Page size: No enforcement
- Title boxes: 3"×1" (Tolerance: 0.2" - more lenient)
- Margins: 1.5" all around (Tolerance: 0.25")
- Logo: Suggested position but AutoFix: No
```

### Configuration C: Icon Library Only
```
- Focus on icon sizes only
- Small icons: 0.5"×0.5" (Tolerance: 0.02")
- Medium icons: 1.0"×1.0" (Tolerance: 0.05")
- Large icons: 2.0"×2.0" (Tolerance: 0.1")
- No position or page validation
```

## Priority Guidelines

Recommended priority order (lowest number = highest priority):

| Priority | Rule Type | Reason |
|----------|-----------|--------|
| 75-79 | Exact positions | Most specific, apply first |
| 80-84 | Page dimensions | Set canvas before positioning |
| 85-89 | Margins | Boundary constraints |
| 90-94 | Shape sizes | Size before final positioning |
| 95-99 | Colors/Fonts | Appearance last |
| 100+ | Text/AI | Content corrections last |

## Tolerance Values Guide

| Shape Type | Recommended Tolerance |
|------------|---------------------|
| Logo placement | 0.05" (very precise) |
| Title boxes | 0.1" (standard) |
| Process boxes | 0.15" (flexible) |
| Decorative elements | 0.25" (very flexible) |
| Page margins | 0.1" (standard) |
| Icon sizes | 0.02" - 0.05" (precise) |

## Troubleshooting

### Shapes Not Resizing

**Problem:** AutoFix is enabled but shapes aren't changing size

**Solutions:**
1. Check shape has text content (only processes visible shapes)
2. Verify vsdx library supports width/height setters
3. Check shape isn't locked or protected
4. Review logs for "Could not set size" warnings

### Position Not Updating

**Problem:** Shapes stay in wrong position despite AutoFix

**Solutions:**
1. Verify x/y properties are accessible
2. Check tolerance isn't too large
3. Ensure shape isn't grouped (grouped shapes may resist)
4. Try ungrouping shapes before upload

### Page Dimensions Not Changing

**Problem:** Pages remain at original size

**Solutions:**
1. Check page.width/height properties exist
2. Verify ExpectedValue format is correct (WIDTHxHEIGHT)
3. Review permission settings (may need page edit rights)
4. Check Visio file isn't protected

## Next Steps

After implementing structural validation:

1. **Test with sample files** before enforcing company-wide
2. **Monitor validation reports** for unexpected issues
3. **Adjust tolerances** based on real-world results
4. **Communicate changes** to diagram creators
5. **Create templates** with pre-validated structure
6. **Document exceptions** for special cases

## Support

For implementation assistance:
- Review main [Visio Validation Guide](visio-validation-guide.md)
- Check Azure Function logs in Application Insights
- Report issues at GitHub

---

**Last Updated:** 2025-01-10
**Version:** 1.0 (Option A Implementation)
