"""
Create a comprehensive test Visio file for structural validation

This file will have violations in:
- Page size (wrong dimensions)
- Shape sizes (various incorrect dimensions)
- Shape positions (outside margins, wrong placement)
- Text content (American spelling, contractions, symbols)
- Fonts (non-Arial if possible)
- Colors (non-standard colors)
"""

import os
import sys

try:
    from vsdx import VisioFile
except ImportError:
    print("Error: vsdx library not found")
    print("Install with: pip install vsdx")
    sys.exit(1)

def create_comprehensive_test_file():
    """Create a Visio file with comprehensive validation violations"""

    print("=" * 70)
    print("CREATING COMPREHENSIVE VISIO TEST FILE")
    print("=" * 70)
    print()

    # Output location
    output_dir = "/Users/stephen/Projects/MaceStyle/test_files"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "test_structural_validation.vsdx")

    print("NOTE: The vsdx library has limitations on creating files from scratch.")
    print("We'll create a template and you'll need to manually add shapes in Visio.")
    print()

    # Create instructions file
    instructions = """
================================================================================
VISIO STRUCTURAL VALIDATION TEST FILE - SETUP INSTRUCTIONS
================================================================================

This test file should demonstrate all structural validation capabilities.

OPEN THE FILE: test_structural_validation.vsdx

Then add the following shapes to create violations:

================================================================================
STEP 1: SET WRONG PAGE SIZE
================================================================================
1. Go to Design → Size → More Page Sizes
2. Set to: 8.5" x 11.0" (Portrait) or 10.0" x 7.5" (Custom)
3. Click OK

EXPECTED FIX: Page will be resized to 11.0" x 8.5" (Letter Landscape)

================================================================================
STEP 2: ADD SHAPES WITH WRONG SIZES
================================================================================

Shape Set A: Title Boxes (should be 3.0" x 1.0")
-------------------------------------------------
1. Insert → Shapes → Rectangle
2. Add text: "Project Title"
3. Resize to: 3.5" x 1.2" (TOO LARGE)
4. Position at: 4.0", 7.5"

5. Insert another rectangle
6. Add text: "Document Title"
7. Resize to: 2.5" x 0.8" (TOO SMALL)
8. Position at: 4.0", 6.5"

Shape Set B: Icons (should be 0.5" x 0.5")
-------------------------------------------
9. Insert → Shapes → Circle
10. Add text: "Icon 1"
11. Resize to: 0.7" x 0.7" (TOO LARGE)
12. Position at: 1.0", 5.0"

13. Insert another circle
14. Add text: "Icon 2"
15. Resize to: 0.3" x 0.3" (TOO SMALL)
16. Position at: 2.0", 5.0"

Shape Set C: Process Boxes (should be 2.0" x 1.5")
---------------------------------------------------
17. Insert → Shapes → Rectangle
18. Add text: "Process Step 1"
19. Resize to: 2.5" x 2.0" (TOO LARGE)
20. Position at: 4.0", 4.0"

21. Insert another rectangle
22. Add text: "Process Step 2"
23. Resize to: 1.5" x 1.0" (TOO SMALL)
24. Position at: 4.0", 2.5"

================================================================================
STEP 3: POSITION SHAPES OUTSIDE MARGINS
================================================================================

Top Margin Violation (should stay below Y=2.0):
-----------------------------------------------
25. Insert rectangle at: 5.0", 2.5" (TOO HIGH)
26. Add text: "Header Text"
27. Size: 2.0" x 0.5"

Left Margin Violation (should start after X=1.0):
--------------------------------------------------
28. Insert rectangle at: 0.5", 4.0" (TOO FAR LEFT)
29. Add text: "Sidebar"
30. Size: 0.5" x 2.0"

Right Margin Violation (should end before X=10.0):
---------------------------------------------------
31. Insert rectangle at: 10.5", 4.0" (TOO FAR RIGHT)
32. Add text: "Note"
33. Size: 0.5" x 1.0"

Bottom Margin Violation (should stay above Y=1.0):
---------------------------------------------------
34. Insert rectangle at: 5.0", 0.5" (TOO LOW)
35. Add text: "Footer Note"
36. Size: 2.0" x 0.3"

================================================================================
STEP 4: TEST EXACT POSITIONING
================================================================================

Logo (should be at 0.5", 7.5"):
--------------------------------
37. Insert rectangle at: 1.0", 7.0" (WRONG POSITION)
38. Add text: "Company Logo"
39. Size: 1.5" x 0.5"
40. Fill with blue color

Footer (should be at 9.5", 0.5"):
----------------------------------
41. Insert rectangle at: 9.0", 1.0" (WRONG POSITION)
42. Add text: "Version 1.0"
43. Size: 1.0" x 0.3"

================================================================================
STEP 5: ADD TEXT WITH STYLE VIOLATIONS
================================================================================

Add these text strings to various shapes (creates AI validation issues):

American Spelling:
- "Project was finalized last week"
- "The color scheme needs authorization"
- "We will analyze the organization structure"
- "Optimization center established"

Contractions:
- "We can't proceed without approval"
- "They won't finalize until it's reviewed"
- "Don't delay the authorization"
- "We're waiting for they've completed"

Symbols:
- "Partnership with M&S"
- "50% complete"
- "Budget increased by 25%"
- "Johnson & Johnson collaboration"

Unformatted Numbers:
- "Budget: 1000 for 5000 items"
- "Processing 10000 records"
- "Cost: 25000 per unit"

Combined Violations:
- "The finalized report can't be authorized until M&S confirms 50% completion for 5000 items"

================================================================================
STEP 6: APPLY WRONG FORMATTING
================================================================================

Fonts (try to vary if possible):
---------------------------------
- Set some shapes to Calibri
- Set some shapes to Times New Roman
- Set some shapes to Verdana
(Note: Visio may enforce consistency, do your best)

Colors:
-------
- Fill some shapes with: Red (#FF0000)
- Fill some shapes with: Green (#00FF00)
- Fill some shapes with: Yellow (#FFFF00)
- Set some text colors to: Red, Blue, Purple

================================================================================
EXPECTED VALIDATION RESULTS
================================================================================

After uploading to SharePoint, the validation should:

✓ Page Dimensions:
  - Resize page from 8.5"×11.0" to 11.0"×8.5"

✓ Shape Sizes:
  - Resize 2 title boxes to 3.0"×1.0"
  - Resize 2 icons to 0.5"×0.5"
  - Resize 2 process boxes to 2.0"×1.5"

✓ Position - Margins:
  - Move header down to Y=2.0" (top margin)
  - Move sidebar right to X=1.0" (left margin)
  - Move note left to X=10.0" (right margin)
  - Move footer up to Y=1.0" (bottom margin)

✓ Position - Exact:
  - Move logo to (0.5", 7.5")
  - Move footer to (9.5", 0.5")

✓ Fonts:
  - Change all text to Arial

✓ Colors:
  - Change fill colors to #003399 (brand blue)
  - Change text colors to #000000 (black)

✓ Text Style (AI):
  - finalized → finalised
  - color → colour
  - analyze → analyse
  - organization → organisation
  - authorization → authorisation
  - can't → cannot
  - won't → will not
  - don't → do not
  - M&S → M and S
  - 50% → 50 percent
  - 1000 → 1,000
  - 5000 → 5,000
  - And many more!

================================================================================
VALIDATION REPORT
================================================================================

The HTML report should show:
- Total issues found: ~50-80 (depending on how many shapes you add)
- Total fixes applied: ~50-80
- Status: PASSED (if all auto-fixed)

Categories:
- Page dimension issues: 1
- Shape size issues: ~6
- Position issues: ~6
- Font issues: ~20-40 (all shapes)
- Color issues: ~10-20 (colored shapes)
- Text style issues: ~20-30 (AI corrections)

================================================================================
SAVE AND UPLOAD
================================================================================

1. Save the file as: test_structural_validation.vsdx
2. Upload to SharePoint document library
3. Wait for validation (10-30 seconds)
4. Download the corrected file
5. Open the validation report
6. Compare before and after!

================================================================================
"""

    # Save instructions
    instructions_path = os.path.join(output_dir, "VISIO_TEST_INSTRUCTIONS.txt")
    with open(instructions_path, "w") as f:
        f.write(instructions)

    print(f"✓ Instructions saved to: {instructions_path}")
    print()

    # Create a simple diagram reference
    print("Creating reference diagram...")

    reference_content = """
VISUAL REFERENCE - Shape Placement Guide
=========================================

Page Layout: 11.0" x 8.5" (Letter Landscape)
Coordinate System: Origin at bottom-left

    Y
    ↑
8.5 ├──────────────────────────────────────────┐ TOP
    │                                          │
7.5 │  [Logo 0.5,7.5]    [Title Center]      │
    │                                          │
6.0 │                                          │
    │                                          │
4.0 │  [Sidebar]      [Process Boxes]         │
    │  X=0.5          X=4.0                   │
2.0 ├─[Header Margin Y=2.0]──────────────────┤
    │                                          │
1.0 ├─[Footer Margin Y=1.0]──────────────────┤
    │                              [Footer]   │
0.0 └──────────────────────────────────────────┘ BOTTOM
    0   1.0              5.0         10.0   11.0  → X
      LEFT            CENTER        RIGHT

Key Zones:
- Top Margin: Y ≤ 2.0" (headers stay below this line)
- Left Margin: X ≥ 1.0" (content starts after this line)
- Right Margin: X ≤ 10.0" (content ends before this line)
- Bottom Margin: Y ≥ 1.0" (footers stay above this line)

Standard Shape Sizes:
- Title Boxes: 3.0" × 1.0"
- Icons: 0.5" × 0.5"
- Process Boxes: 2.0" × 1.5"

Exact Positions:
- Logo: (0.5", 7.5") - top-left corner
- Footer: (9.5", 0.5") - bottom-right corner
"""

    reference_path = os.path.join(output_dir, "VISIO_LAYOUT_REFERENCE.txt")
    with open(reference_path, "w") as f:
        f.write(reference_content)

    print(f"✓ Layout reference saved to: {reference_path}")
    print()

    # Create a quick setup checklist
    checklist = """
QUICK SETUP CHECKLIST
=====================

□ Open Microsoft Visio Desktop (required, not Visio Online)
□ Create new blank drawing
□ Change page size to 8.5" × 11.0" portrait (wrong size)

Add Shapes with Violations:
□ 2 title boxes (wrong sizes: 3.5"×1.2" and 2.5"×0.8")
□ 2 icons (wrong sizes: 0.7"×0.7" and 0.3"×0.3")
□ 2 process boxes (wrong sizes: 2.5"×2.0" and 1.5"×1.0")
□ 1 header (wrong position: Y=2.5")
□ 1 sidebar (wrong position: X=0.5")
□ 1 right note (wrong position: X=10.5")
□ 1 footer note (wrong position: Y=0.5")
□ 1 logo (wrong position: X=1.0", Y=7.0")
□ 1 version footer (wrong position: X=9.0", Y=1.0")

Add Text with Style Violations:
□ American spelling (finalized, color, analyze, etc.)
□ Contractions (can't, won't, don't, etc.)
□ Symbols (M&S, 50%, etc.)
□ Unformatted numbers (1000, 5000, etc.)

Apply Wrong Formatting:
□ Set various fonts (Calibri, Times New Roman, etc.)
□ Apply various colors (red, green, yellow fills)
□ Apply colored text (red, blue, purple)

Save and Test:
□ Save as: test_structural_validation.vsdx
□ Upload to SharePoint document library
□ Wait for validation
□ Download corrected file
□ View validation report

Expected Results:
□ Page resized to 11.0"×8.5"
□ All shapes resized to standards
□ All shapes repositioned to margins/exact positions
□ All fonts changed to Arial
□ All colors changed to brand standards
□ All text corrected for British English
□ Validation report shows all fixes
"""

    checklist_path = os.path.join(output_dir, "SETUP_CHECKLIST.txt")
    with open(checklist_path, "w") as f:
        f.write(checklist)

    print(f"✓ Setup checklist saved to: {checklist_path}")
    print()

    print("=" * 70)
    print("FILES CREATED")
    print("=" * 70)
    print()
    print(f"1. Instructions:    {instructions_path}")
    print(f"2. Layout Guide:    {reference_path}")
    print(f"3. Quick Checklist: {checklist_path}")
    print()
    print("=" * 70)
    print("NEXT STEPS")
    print("=" * 70)
    print()
    print("The vsdx Python library cannot create shapes from scratch.")
    print("You need to manually create the test file in Microsoft Visio.")
    print()
    print("Follow these steps:")
    print("1. Open Microsoft Visio Desktop")
    print("2. Create a new blank drawing")
    print("3. Follow the instructions in: VISIO_TEST_INSTRUCTIONS.txt")
    print("4. Save as: test_structural_validation.vsdx")
    print("5. Upload to SharePoint to test validation")
    print()
    print("OR for a quick test:")
    print("- Create any Visio file with wrong page size and shapes")
    print("- Add text with American spelling")
    print("- Upload to SharePoint")
    print("- The validation will still work!")
    print()
    print("=" * 70)

    return True

if __name__ == "__main__":
    try:
        success = create_comprehensive_test_file()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
