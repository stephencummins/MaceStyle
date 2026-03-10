"""Local testing without SharePoint connection"""
import sys
import json
from io import BytesIO
from ValidateDocument.word_validator import validate_word_document
from ValidateDocument.report import generate_report
from ValidateDocument.test_helpers import create_test_word_doc, create_mock_rules

def test_validation():
    print("🧪 Testing Word validation locally...")
    
    # Create test document
    print("\n1. Creating test Word document with issues...")
    test_doc_stream = create_test_word_doc()
    
    # Load mock rules
    print("2. Loading mock validation rules...")
    rules = create_mock_rules()
    print(f"   Loaded {len(rules)} rules")
    
    # Run validation
    print("\n3. Running validation...")
    result = validate_word_document(test_doc_stream, rules)
    
    # Display results
    print(f"\n✅ Validation complete!")
    print(f"   Issues found: {len(result['issues'])}")
    print(f"   Fixes applied: {len(result['fixes_applied'])}")
    
    if result['issues']:
        print("\n📋 Issues:")
        for issue in result['issues']:
            print(f"   - {issue}")
    
    if result['fixes_applied']:
        print("\n🔧 Fixes:")
        for fix in result['fixes_applied']:
            print(f"   - {fix}")
    
    # Generate report
    print("\n4. Generating report...")
    report = generate_report("test.docx", result['issues'], result['fixes_applied'])
    print(f"   Report length: {len(report)} characters")
    
    # Save fixed document
    print("\n5. Saving fixed document...")
    fixed_stream = BytesIO()
    result['document'].save(fixed_stream)
    fixed_stream.seek(0)
    
    with open('test_output_fixed.docx', 'wb') as f:
        f.write(fixed_stream.getvalue())
    print("   ✅ Saved to: test_output_fixed.docx")
    
    # Save report
    with open('test_output_report.html', 'w') as f:
        f.write(report)
    print("   ✅ Saved to: test_output_report.html")
    
    print("\n✅ Local test complete! Check the output files.")

if __name__ == "__main__":
    test_validation()