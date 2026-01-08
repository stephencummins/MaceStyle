# Process Map Validation Rules - Summary

## ✅ Successfully Added: 22 Rules

All process map validation rules from the PowerPoint template have been added to SharePoint.

## 📊 Rule Categories

### 1. Page Structure (1 rule)
- **Slide Size 10×5.62 inches**
  - RuleType: PageDimensions
  - AutoFix: Yes
  - Ensures all process maps use widescreen format

### 2. Required Content (5 rules)
- **Top Line Workstream Name Required**
  - Must display workstream name at top of map
  - AI-powered validation

- **Bottom Line Process Name Required**
  - Must display process name at bottom of map
  - AI-powered validation

- **Document Reference Required**
  - Doc Ref: [value] must be present
  - Assigned by Content/Document Management Team

- **Set Up or Start Activity Required**
  - Every map must begin with "Set Up" or "Start" activity
  - AI-powered validation

- **Feedback Activity Required**
  - Must include feedback activity at end
  - Must loop back to Set Up/Start
  - AI-powered validation

### 3. Swim Lanes (6 rules)

**Required Structure:**
- All 5 swim lanes must be present on every map:
  1. New Hospitals Programme
  2. NHS
  3. Healthy Delivery Partnership
  4. Delivery Team
  5. Contractor/Supply Chain

**Color Rules:**
- Each swim lane has color validation
- Boxes in swim lane must use lane color in header
- AI-powered validation for compliance

### 4. Box Formatting (4 rules)

- **Box Headers - One Word or Short Sentence**
  - Each box header must be concise
  - Summary of the activity only

- **Box Body - Concise Description**
  - Activity/task description supported by Activity Guide
  - Must NOT specify role ownership
  - AI-powered validation

- **Multi-Lane Activities - Grey Color**
  - Activities spanning >1 swim lane = grey boxes
  - Color: #808080 (grey)
  - AutoFix: Yes

- **Interface Boxes - Dark Grey**
  - Located at bottom of process steps
  - Link to other processes
  - Color: #404040 (dark grey)
  - AutoFix: Yes

### 5. Stage Gates (1 rule)

- **RIBA Plan of Work Format**
  - Stage Gates must follow RIBA standards
  - Stages not applicable should be deleted
  - AI-powered validation

### 6. Decision Points (2 rules)

- **Diamond Shapes Required**
  - Decision points must use diamond shapes
  - AI-powered validation (Visio only)

- **YES/NO Labels Required**
  - Arrows from decision points must contain explicit options
  - Must say "YES" or "NO"
  - AI-powered validation

### 7. Activity Guides (2 rules)

- **Underlined Headers Need Activity Guide**
  - Underlined box headers = single Activity Guide required
  - Non-underlined = covered by general Activity Guide

- **Activity Guide References Included**
  - References must be included where applicable
  - AI-powered validation

### 8. Activity References (1 rule)

- **Format: XX-000001**
  - Activity references must follow this format
  - AI-powered validation

## 🎯 Validation Approach

### AI-Powered Rules (18 rules)
These rules use Claude AI to validate complex requirements:
- Content requirements (workstream name, process name, etc.)
- Structural requirements (swim lanes, activities, feedback loop)
- Text formatting (box headers, descriptions)
- Decision point labels
- Activity guide references

**Cannot auto-fix** - AI will detect violations and report them

### Hard-Coded Rules (4 rules)
These rules use direct validation:
- Page dimensions (10.0×5.62)
- Multi-lane activity color (#808080)
- Interface box color (#404040)

**Can auto-fix** - Automatically corrects violations

## 📝 Rule Priority Levels

| Priority Range | Category |
|---------------|----------|
| 50 | Page structure (highest) |
| 110-114 | Required content |
| 115 | Swim lanes |
| 116-118 | Stage gates and box formatting |
| 119-120 | Color rules |
| 121-125 | Swim lane colors |
| 126-130 | Decision points, activity guides, references |

Lower numbers = higher priority = validated first

## 🚀 How Validation Works

When a process map is uploaded to SharePoint:

1. **Structural Checks:**
   - Page size → Auto-fix to 10.0×5.62
   - Grey/dark grey boxes → Auto-fix colors

2. **AI-Powered Checks (Claude):**
   - Scans entire map for required elements
   - Validates swim lane presence
   - Checks box header/body formatting
   - Verifies decision point labels
   - Confirms activity references format
   - Validates Stage Gates format

3. **Report Generation:**
   - Lists all violations found
   - Shows which were auto-fixed
   - Flags violations requiring manual correction

## ⚠️ Important Notes

### Cannot Auto-Fix:
- Missing required content (workstream name, process name, etc.)
- Missing swim lanes
- Incorrect box header format
- Missing decision point labels
- Incorrect activity reference format

These will be **flagged in the report** but require manual correction.

### Can Auto-Fix:
- Wrong page dimensions
- Incorrect grey box colors
- Incorrect dark grey interface box colors

## 📋 Complete Rule List

| # | Rule Title | Type | AutoFix | UseAI |
|---|-----------|------|---------|-------|
| 1 | Slide Size 10x5.62 inches | PageDimensions | Yes | No |
| 2 | Top Line Workstream Name Required | Layout | No | Yes |
| 3 | Bottom Line Process Name Required | Layout | No | Yes |
| 4 | Document Reference Required | Layout | No | Yes |
| 5 | Set Up or Start Activity Required | Layout | No | Yes |
| 6 | Feedback Activity Required | Layout | No | Yes |
| 7 | All 5 Swim Lanes Required | Layout | No | Yes |
| 8 | RIBA Stage Gates Format | Language | No | Yes |
| 9 | Box Headers One Word or Short Sentence | Grammar | No | Yes |
| 10 | Box Body Concise Description | Grammar | No | Yes |
| 11 | Multi-Lane Activities Grey Color | Color | Yes | No |
| 12 | Interface Boxes Dark Grey | Color | Yes | No |
| 13 | New Hospitals Programme Lane Color | Color | No | Yes |
| 14 | NHS Lane Color | Color | No | Yes |
| 15 | Healthy Delivery Partnership Lane Color | Color | No | Yes |
| 16 | Delivery Team Lane Color | Color | No | Yes |
| 17 | Contractor Lane Color | Color | No | Yes |
| 18 | Decision Points Use Diamond Shapes | Layout | No | Yes |
| 19 | Decision Arrows Have YES/NO Labels | Language | No | Yes |
| 20 | Underlined Headers Need Activity Guide | Layout | No | Yes |
| 21 | Activity Guide References Included | Language | No | Yes |
| 22 | Activity Reference Format XX-000001 | Language | No | Yes |

## 🔧 Testing

To test these rules:

1. **Create a process map** using the template
2. **Intentionally violate rules:**
   - Use wrong page size
   - Miss required elements (no "Start" activity)
   - Use wrong colors for interface boxes
   - Omit swim lanes
   - Don't include document reference
3. **Upload to SharePoint**
4. **Review validation report**

Expected violations to be flagged:
- Missing workstream/process names
- Missing swim lanes
- Missing Start/feedback activities
- Incorrect colors (will be auto-fixed)
- Wrong page size (will be auto-fixed)

## 📖 Reference Documents

- **Original Rules:** From PowerPoint template guidance page
- **Implementation:** `MaceStyleValidator/add_process_map_rules.py`
- **SharePoint List:** Style Rules list in SharePoint site

## 🎉 Next Steps

1. ✓ Rules added to SharePoint
2. Test with sample process map
3. Review validation reports
4. Refine rules based on results
5. Train team on requirements
6. Integrate into workflow

## 💡 Tips for Process Map Creators

To ensure your process maps pass validation:

- ✓ Start with the template
- ✓ Include all 5 swim lanes (even if empty)
- ✓ Add workstream name at top, process name at bottom
- ✓ Include "Set Up" or "Start" activity
- ✓ Add feedback loop at end
- ✓ Use diamond shapes for decisions
- ✓ Label decision arrows with YES/NO
- ✓ Keep box headers short (one word or brief sentence)
- ✓ Don't specify roles in box descriptions
- ✓ Add document reference (get from Content Management Team)
- ✓ Use grey for multi-lane activities
- ✓ Use dark grey for interface boxes
- ✓ Follow RIBA Plan of Work for Stage Gates
- ✓ Add Activity Guide references where needed

---

**Date Added:** 2025-01-10
**Total Rules:** 22
**Status:** ✅ Active in SharePoint
**Applies To:** Visio and PowerPoint process maps
