"""
Create a test Visio file with content to validate against style rules
"""
from vsdx import VisioFile
import os

def create_test_visio():
    """Create a test Visio diagram with text containing style violations"""

    print("Creating test Visio file with style violations...\n")

    # Create a new blank Visio file
    # Note: We'll start with a basic template
    output_dir = "/Users/stephen/Projects/MaceStyle/test_files"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "test_visio_validation.vsdx")

    # Create a new Visio file
    vis = VisioFile()

    # Get the first page
    if len(vis.pages) > 0:
        page = vis.pages[0]
        page.name = "Style Validation Test"

        print("✅ Created Visio file with page: 'Style Validation Test'")
        print("\nThe file has been created, but requires manual setup.")
        print("\n" + "="*70)
        print("IMPORTANT: Manual setup required")
        print("="*70)
        print("\nThe vsdx Python library has limited shape creation capabilities.")
        print("Please follow these steps to create a testable Visio diagram:\n")

        print("1. Save the generated file")
        vis.save_vsdx(output_path)
        print(f"   ✓ Saved to: {output_path}\n")

        print("2. Open in Microsoft Visio Desktop")
        print("   (Visio Online has limited text editing)\n")

        print("3. Add the following shapes with text:\n")

        test_content = [
            ("Rectangle 1", "Project Status - finalized", "Should be 'finalised'"),
            ("Rectangle 2", "We can't proceed without approval", "Should be 'cannot'"),
            ("Rectangle 3", "The color scheme is complete", "Should be 'colour'"),
            ("Rectangle 4", "M&S Partnership - 50% growth", "Should be 'M and S' and '50 percent'"),
            ("Rectangle 5", "Analysis Center - Organization", "Should be 'Centre' and 'Organisation'"),
            ("Rectangle 6", "Don't delay the authorization", "Should be 'Do not' and 'authorisation'"),
            ("Rectangle 7", "Budget: 1000 for 5000 items", "Should be '1,000' and '5,000'"),
            ("Rectangle 8", "We won't finalize until they've analyzed the data", "Multiple violations"),
        ]

        for i, (shape, text, note) in enumerate(test_content, 1):
            print(f"   Shape {i}: {shape}")
            print(f"      Text: \"{text}\"")
            print(f"      Note: {note}")
            print()

        print("4. Format some shapes with non-standard fonts:")
        print("   - Set some text to Calibri")
        print("   - Set some text to Times New Roman")
        print("   - Leave some as default (should be Arial)\n")

        print("5. Save the file\n")

        print("6. Upload to SharePoint to test validation\n")

        print("="*70)
        print("Expected Corrections:")
        print("="*70)
        print("✓ finalized → finalised")
        print("✓ can't → cannot, don't → do not, won't → will not")
        print("✓ color → colour, center → centre, organization → organisation")
        print("✓ M&S → M and S")
        print("✓ 50% → 50 percent")
        print("✓ 1000 → 1,000, 5000 → 5,000")
        print("✓ All fonts → Arial (if font validation is enabled for Visio)")
        print("\n" + "="*70)
        print("\n⚠️  NOTE: Visio validation is currently MINIMAL in the code.")
        print("The main validation engine is designed for Word documents.")
        print("You may need to enhance the Visio validation logic in:")
        print("  - MaceStyleValidator/ValidateDocument/__init__.py (lines 418-469)")
        print("\n" + "="*70)

    else:
        print("❌ Could not create Visio page")
        return None

    # Create a detailed reference document
    reference_content = """
VISIO TEST FILE - MANUAL SETUP GUIDE
====================================

File Location: {output_path}

STEP-BY-STEP INSTRUCTIONS:

1. OPEN IN VISIO DESKTOP
   - Double-click the file or open in Microsoft Visio
   - Visio Online may have limited capabilities

2. ADD TEST SHAPES
   Create 8 rectangles with the following text:

   Shape 1: "Project Status - finalized"
   Shape 2: "We can't proceed without approval"
   Shape 3: "The color scheme is complete"
   Shape 4: "M&S Partnership - 50% growth"
   Shape 5: "Analysis Center - Organization"
   Shape 6: "Don't delay the authorization"
   Shape 7: "Budget: 1000 for 5000 items"
   Shape 8: "We won't finalize until they've analyzed the data"

3. FORMAT VARIATIONS (Optional)
   - Set Shape 1-2: Calibri font
   - Set Shape 3-4: Times New Roman
   - Set Shape 5-6: Arial (correct)
   - Leave Shape 7-8: Default

4. ADD DIAGRAM STRUCTURE (Optional)
   - Connect shapes with arrows
   - Add a title: "Process Flow Diagram"
   - Add colors and styling

5. SAVE THE FILE
   File → Save

6. TEST VALIDATION
   - Upload to SharePoint Document Library
   - Wait for validation
   - Check results

EXPECTED CORRECTIONS:
====================

British English:
- finalized → finalised
- color → colour
- center → centre
- organization → organisation
- authorization → authorisation
- analyzed → analysed

Grammar (Contractions):
- can't → cannot
- don't → do not
- won't → will not
- they've → they have

Symbols:
- M&S → M and S
- 50% → 50 percent

Numbers:
- 1000 → 1,000
- 5000 → 5,000

Fonts (if enabled):
- All text → Arial

CURRENT LIMITATIONS:
===================

The Visio validation logic in the Azure Function is currently minimal:
- Basic structure exists (lines 418-469 in __init__.py)
- check_visio_colors() and check_visio_fonts() are stubs
- No text validation implemented yet

TO ENABLE FULL VISIO VALIDATION:
================================

You would need to enhance:

1. Text extraction from Visio shapes
2. Apply the same AI validation as Word docs
3. Text replacement in Visio shapes
4. Font and color validation for Visio

The vsdx Python library supports:
- Reading shape text
- Modifying shape text
- Accessing shape properties

Example enhancement location:
  MaceStyleValidator/ValidateDocument/__init__.py
  Function: validate_visio_document()

For now, the file serves as a test artifact to verify:
- Power Automate triggers on Visio files
- Azure Function receives Visio files
- Error handling for unsupported operations
""".format(output_path=output_path)

    reference_path = os.path.join(output_dir, "VISIO_TEST_SETUP_GUIDE.txt")
    with open(reference_path, "w") as f:
        f.write(reference_content)

    print(f"\n📝 Setup guide saved to: {reference_path}")

    return output_path

if __name__ == "__main__":
    try:
        path = create_test_visio()
        if path:
            print("\n" + "="*70)
            print("✅ SUCCESS - Visio template created")
            print("="*70)
            print(f"\n📂 Next: Open {path} in Microsoft Visio")
            print("📋 Then: Follow the setup guide to add test content")
            print("🚀 Finally: Upload to SharePoint to test validation")
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()
