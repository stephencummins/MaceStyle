# Visio Test File Creation Guide

## ⚠️ Important Note

**Visio validation is currently MINIMAL** in the MaceStyle Validator. The system is primarily designed for Word document validation. This guide helps you create a test Visio file to verify:
- Power Automate triggers correctly on Visio files
- Azure Function receives and processes Visio files
- Error handling for unsupported operations

## 📋 How to Create a Test Visio File

### Step 1: Create New Visio Document

1. Open **Microsoft Visio Desktop** (not Visio Online)
2. Create a new **Blank Drawing**
3. Save as: `test_visio_validation.vsdx`

### Step 2: Add Test Shapes with Style Violations

Create the following shapes (rectangles or process boxes) with text:

| Shape | Text Content | Expected Corrections |
|-------|-------------|---------------------|
| **1** | `Project Status - finalized` | finalised |
| **2** | `We can't proceed without approval` | cannot |
| **3** | `The color scheme is complete` | colour |
| **4** | `M&S Partnership - 50% growth` | M and S, 50 percent |
| **5** | `Analysis Center - Organization` | Centre, Organisation |
| **6** | `Don't delay the authorization` | Do not, authorisation |
| **7** | `Budget: 1000 for 5000 items` | 1,000 and 5,000 |
| **8** | `We won't finalize until they've analyzed` | Multiple corrections |

### Step 3: Add Font Variations (Optional)

- **Shapes 1-2**: Set to Calibri font
- **Shapes 3-4**: Set to Times New Roman
- **Shapes 5-6**: Set to Arial (correct)
- **Shapes 7-8**: Leave as default

### Step 4: Create a Simple Diagram (Optional)

1. Connect shapes with arrows
2. Add a title: "Process Flow Diagram"
3. Add some colors and styling
4. Group related shapes

### Step 5: Save the File

- **File** → **Save**
- Location: `test_files/test_visio_validation.vsdx`

### Step 6: Upload to SharePoint

1. Go to your SharePoint Document Library
2. Upload `test_visio_validation.vsdx`
3. Watch for validation trigger
4. Check results

## 🎯 Expected Results

### Current Behavior (Minimal Validation)

Since Visio validation is minimal, the function will:
- ✅ Receive the file successfully
- ✅ Identify it as a Visio file (`.vsdx`)
- ⚠️ Apply limited validation (structure exists but not fully implemented)
- ⚠️ May return "Passed" with no changes

### What SHOULD Happen (If Fully Implemented)

The same corrections as Word documents:

#### British English Spelling
- finalized → finalised
- color → colour
- center → centre
- organization → organisation
- authorization → authorisation
- analyzed → analysed

#### Grammar (Contractions)
- can't → cannot
- don't → do not
- won't → will not
- they've → they have

#### Symbols
- M&S → M and S
- 50% → 50 percent

#### Numbers
- 1000 → 1,000
- 5000 → 5,000

#### Fonts (if enabled)
- All text → Arial

## 🔧 Current Visio Validation Implementation

### Code Location
`MaceStyleValidator/ValidateDocument/__init__.py` (lines 418-469)

### What Exists
```python
def validate_visio_document(file_stream, rules):
    """Validate Visio document against rules"""
    visio = VisioFile(file_stream)
    issues = []
    fixes_applied = []

    # Filter rules for Visio documents
    visio_rules = [r for r in rules if r['doc_type'] == 'Visio']

    for rule in visio_rules:
        if rule['rule_type'] == 'Color':
            result = check_visio_colors(visio, rule)
            # ...
        elif rule['rule_type'] == 'Font':
            result = check_visio_fonts(visio, rule)
            # ...

    return {
        'document': visio,
        'issues': issues,
        'fixes_applied': fixes_applied
    }
```

### What's Missing
- ❌ Text extraction from shapes
- ❌ AI validation for text content
- ❌ Text replacement in shapes
- ❌ Comprehensive font validation
- ❌ Color validation implementation

## 🚀 How to Enhance Visio Validation

If you want to add full Visio support, you would need to:

### 1. Extract Text from All Shapes

```python
def extract_visio_text(visio):
    """Extract all text from Visio shapes"""
    all_text = []
    for page in visio.pages:
        for shape in page.shapes:
            if shape.text:
                all_text.append({
                    'shape_id': shape.ID,
                    'text': shape.text
                })
    return all_text
```

### 2. Apply AI Validation

Use the same Claude AI validation as Word documents:

```python
# Extract all text
shape_texts = extract_visio_text(visio)
combined_text = "\n".join([s['text'] for s in shape_texts])

# Send to Claude AI
corrections = validate_with_claude(combined_text, ai_rules)

# Apply corrections back to shapes
for shape_data, correction in zip(shape_texts, corrections):
    shape = find_shape_by_id(visio, shape_data['shape_id'])
    shape.text = correction
```

### 3. Font and Color Validation

```python
def check_visio_fonts(visio, rule):
    """Check and fix fonts in Visio diagrams"""
    issues = []
    fixes = []
    expected_font = rule['expected_value']

    for page in visio.pages:
        for shape in page.shapes:
            # Access shape font properties
            # Note: vsdx library has methods for this
            if shape.has_text():
                current_font = shape.get_font()
                if current_font != expected_font:
                    if rule['auto_fix']:
                        shape.set_font(expected_font)
                        fixes.append(f"Fixed font in shape {shape.ID}")
                    else:
                        issues.append(f"Wrong font in shape {shape.ID}")

    return {'issues': issues, 'fixes': fixes}
```

## 📚 Resources

- **vsdx Library Documentation**: https://pypi.org/project/vsdx/
- **Visio File Format**: https://docs.microsoft.com/en-us/office/client-developer/visio/introduction-to-the-visio-file-formatvsdx

## 💡 Alternative Approach

If full Visio validation is important, consider:

1. **Export to PDF** - Extract text from PDF for validation
2. **OCR Processing** - Use Azure Cognitive Services to extract text
3. **Manual Review** - Keep Visio validation manual with guidelines
4. **Word Export** - Have users export diagrams to Word for validation

## ✅ Quick Test Checklist

- [ ] Created Visio file with test content
- [ ] Added 8 shapes with style violations
- [ ] Saved file as `test_visio_validation.vsdx`
- [ ] Uploaded to SharePoint Document Library
- [ ] Power Automate flow triggered
- [ ] Azure Function received file
- [ ] Validation completed (even if minimal)
- [ ] Reviewed results in Validation Results list

## 🎯 Success Criteria

For this test, success means:
- ✅ File uploads without errors
- ✅ Validation status updates
- ✅ No crashes or exceptions
- ✅ Some result returned (even if "Passed" with no changes)

Full text validation would be a future enhancement!
