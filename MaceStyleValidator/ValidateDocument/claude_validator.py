"""
AI-powered document validation using Claude Haiku
"""
import os
import logging
import json
from anthropic import Anthropic

def extract_document_text(doc):
    """Extract all text from Word document"""
    text_parts = []

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            text_parts.append(paragraph.text)

    return "\n\n".join(text_parts)

def build_validation_prompt(document_text, ai_rules):
    """Build comprehensive prompt for Claude with all AI-enabled rules"""

    # Group rules by type for better organization
    rules_by_type = {}
    for rule in ai_rules:
        rule_type = rule.get('rule_type', 'Other')
        if rule_type not in rules_by_type:
            rules_by_type[rule_type] = []
        rules_by_type[rule_type].append(rule)

    # Build rules description
    rules_description = []
    for rule_type, rules in rules_by_type.items():
        rules_description.append(f"\n**{rule_type} Rules:**")
        for rule in rules:
            rules_description.append(f"- {rule.get('title', 'Unknown rule')}")

    prompt = f"""You are a professional document editor specializing in the Mace Control Centre Writing Style Guide.

Your task is to review and correct the following document according to these style rules:
{''.join(rules_description)}

**Key Requirements:**
1. Convert ALL American spellings to British English (colour, centre, analyse, organisation, etc.)
2. Expand ALL contractions in formal text (can't → cannot, don't → do not, etc.)
3. Replace ampersands (&) with 'and'
4. Replace percent symbols (%) with the word 'percent'
5. Use 'toward' not 'towards'
6. Avoid 'etc.' - be specific instead
7. Add commas to numbers 1000+ (e.g., 1,000; 15,500)

**Document Text:**
{document_text}

**Instructions:**
1. Return the CORRECTED document text
2. List ALL changes you made in this format:
   - Issue: [describe what was wrong]
   - Fix: [describe what you changed]

Return your response as JSON:
{{
  "corrected_text": "the full corrected document text here",
  "changes": [
    {{"issue": "description of issue", "fix": "description of fix"}},
    ...
  ]
}}"""

    return prompt

def apply_corrections_to_document(doc, corrected_text):
    """Apply Claude's corrections back to the Word document"""
    import logging

    logging.info("Applying corrections to document...")

    # Split corrected text into paragraphs
    corrected_paragraphs = corrected_text.split('\n\n')

    # Match to original paragraphs and update
    doc_para_index = 0
    updated_count = 0

    for i, corrected_para in enumerate(corrected_paragraphs):
        if not corrected_para.strip():
            continue

        # Find next non-empty paragraph in document
        while doc_para_index < len(doc.paragraphs):
            if doc.paragraphs[doc_para_index].text.strip():
                break
            doc_para_index += 1

        if doc_para_index >= len(doc.paragraphs):
            logging.warning(f"Ran out of document paragraphs at index {i}")
            break

        # Update the paragraph text while preserving formatting
        paragraph = doc.paragraphs[doc_para_index]

        if len(paragraph.runs) > 0:
            # Clear existing runs and create one new run with corrected text
            for run in paragraph.runs[1:]:
                run.text = ""
            paragraph.runs[0].text = corrected_para.strip()
            updated_count += 1
        else:
            # No runs exist, add the text directly
            paragraph.text = corrected_para.strip()
            updated_count += 1

        doc_para_index += 1

    logging.info(f"Applied corrections to {updated_count} paragraphs")

def validate_with_claude(doc, ai_rules):
    """
    Validate document using Claude Haiku API

    Args:
        doc: python-docx Document object
        ai_rules: List of rules marked with UseAI=True

    Returns:
        dict with 'issues' and 'fixes_applied' lists
    """
    logging.info("=" * 60)
    logging.info("CLAUDE VALIDATOR v2.1 - SIMPLE TEST")
    logging.info("=" * 60)

    issues = []
    fixes_applied = []

    # SIMPLE TEST: Just modify the first paragraph
    logging.info("SIMPLE TEST: Modifying first paragraph...")
    if len(doc.paragraphs) > 0:
        first_para = doc.paragraphs[0]
        original_text = first_para.text
        logging.info(f"Original first para: {original_text[:100]}")

        if len(first_para.runs) > 0:
            # Change "color" to "COLOUR" as a test
            new_text = first_para.runs[0].text.replace("color", "COLOUR").replace("Color", "COLOUR")
            first_para.runs[0].text = new_text
            logging.info(f"Modified first para: {new_text[:100]}")
            issues.append("Test: Found 'color'")
            fixes_applied.append("Test: Changed to 'COLOUR'")

    logging.info(f"Simple test complete: {len(issues)} issues, {len(fixes_applied)} fixes")

    return {
        'issues': issues,
        'fixes_applied': fixes_applied
    }

    # Check if Anthropic API key is configured
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        error_msg = "ANTHROPIC_API_KEY not configured - skipping AI validation"
        logging.error(error_msg)
        return {'issues': [error_msg], 'fixes_applied': []}

    try:
        logging.info(f"Starting Claude AI validation with {len(ai_rules)} rules")
        logging.info(f"AI Rules received: {[r.get('title', 'Unknown') for r in ai_rules]}")

        # Extract document text
        document_text = extract_document_text(doc)
        if not document_text.strip():
            logging.warning("Document has no text to validate")
            return {'issues': [], 'fixes_applied': []}

        logging.info(f"Extracted {len(document_text)} characters from document")

        # Build prompt
        prompt = build_validation_prompt(document_text, ai_rules)

        # Call Claude Haiku
        client = Anthropic(api_key=api_key)

        logging.info("Calling Claude Haiku API...")
        logging.info(f"Prompt length: {len(prompt)} characters")

        try:
            response = client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=4096,  # Claude Haiku max output tokens
                temperature=0.3,  # Lower temperature for consistency
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
        except Exception as api_error:
            logging.error(f"Claude API error: {str(api_error)}")
            logging.error(f"Error type: {type(api_error).__name__}")
            if hasattr(api_error, 'response'):
                logging.error(f"API response: {api_error.response}")
            raise

        # Parse response
        response_text = response.content[0].text
        logging.info(f"Received response from Claude ({len(response_text)} characters)")

        # Extract JSON from response (Claude sometimes wraps it in markdown)
        json_start = response_text.find('{')
        json_end = response_text.rfind('}') + 1
        if json_start >= 0 and json_end > json_start:
            json_text = response_text[json_start:json_end]

            logging.info(f"Parsing JSON response ({len(json_text)} chars)")

            # Use strict=False to allow control characters
            try:
                result = json.loads(json_text, strict=False)
            except json.JSONDecodeError as e:
                logging.error(f"JSON decode error: {e}")
                # If strict=False doesn't work, try escaping newlines/tabs
                import re
                # Properly escape newlines, tabs, and carriage returns in JSON string values
                json_text_escaped = json_text.replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')
                logging.info("Retrying with escaped control characters")
                result = json.loads(json_text_escaped)
        else:
            raise ValueError("Could not find JSON in Claude's response")

        # Skip verbose inspection - just extract data
        logging.info("Extracting corrected_text and changes from result...")

        try:
            corrected_text = result.get('corrected_text', '')
            changes = result.get('changes', [])

            logging.info(f"Extracted: {len(corrected_text)} chars, {len(changes)} changes")

            # Apply corrections to document
            if corrected_text and changes:
                logging.info("Applying corrections to document...")
                apply_corrections_to_document(doc, corrected_text)

                # Build issues and fixes lists
                for change in changes:
                    issue_desc = change.get('issue', 'Unknown issue')
                    fix_desc = change.get('fix', 'Applied correction')
                    issues.append(issue_desc)
                    fixes_applied.append(fix_desc)

                logging.info(f"Applied {len(fixes_applied)} fixes")
            else:
                logging.warning(f"Skipping - no text or changes")
        except Exception as apply_error:
            import traceback
            logging.error(f"Error in correction workflow: {str(apply_error)}")
            logging.error(f"Error type: {type(apply_error).__name__}")
            logging.error(f"Traceback: {traceback.format_exc()}")
            raise

        logging.info("=" * 60)
        logging.info(f"RETURNING: issues={len(issues)}, fixes_applied={len(fixes_applied)}")
        logging.info("=" * 60)
        return {
            'issues': issues,
            'fixes_applied': fixes_applied
        }

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        logging.error(f"Claude AI validation failed: {str(e)}")
        logging.error(f"Full traceback: {error_details}")
        # Return error but don't fail the entire validation
        return {
            'issues': [f"AI validation error: {str(e)}"],
            'fixes_applied': []
        }
