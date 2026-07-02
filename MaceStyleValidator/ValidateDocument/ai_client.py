"""Centralised Claude AI client for style validation"""
import os
import json
import logging
from anthropic import Anthropic
from .config import ENABLE_CLAUDE_AI, AI_PROVIDER, CLAUDE_MODEL, CLAUDE_MAX_TOKENS, CLAUDE_TEMPERATURE


def get_ai_client():
    """Return a Messages-API client for the configured provider, or None if unconfigured.

    Both providers expose the identical client.messages.create() surface, so the
    rest of this module is provider-agnostic. Switching to Mace's Foundry later
    is an app-settings change only (AI_PROVIDER, FOUNDRY_RESOURCE, FOUNDRY_API_KEY,
    CLAUDE_MODEL) - no code change.
    """
    if AI_PROVIDER == "foundry":
        from anthropic import AnthropicFoundry
        resource = os.environ.get("FOUNDRY_RESOURCE")
        api_key = os.environ.get("FOUNDRY_API_KEY")
        if not (resource and api_key):
            logging.warning("FOUNDRY_RESOURCE/FOUNDRY_API_KEY not set - skipping AI validation")
            return None
        return AnthropicFoundry(resource=resource, api_key=api_key)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        logging.warning("ANTHROPIC_API_KEY not set - skipping AI validation")
        return None
    return Anthropic(api_key=api_key)


def build_dynamic_prompt(ai_rules, document_text):
    """Build Claude prompt dynamically from rules where UseAI=True"""
    rules_by_type = {}
    for rule in ai_rules:
        rule_type = rule.get('rule_type', 'Other')
        if rule_type not in rules_by_type:
            rules_by_type[rule_type] = []
        rules_by_type[rule_type].append(rule)

    rules_description = []
    for rule_type, rules in sorted(rules_by_type.items()):
        rules_description.append(f"\n**{rule_type} Rules:**")
        for rule in rules:
            title = rule.get('title', 'Unknown rule')
            expected = rule.get('expected_value', '')
            if expected:
                rules_description.append(f"- {title} (use: {expected})")
            else:
                rules_description.append(f"- {title}")

    return f"""You are a professional document editor applying the Mace Control Centre Writing Style Guide.

Apply ALL of the following corrections to the text:
{''.join(rules_description)}

Return a JSON object with two fields:
1. "corrected_text": the full corrected text (preserve paragraph breaks as \\n\\n)
2. "changes_made": total count of ALL changes made

Text to correct:
{document_text}"""


def call_claude(ai_rules, document_text):
    """Call Claude API for style validation.

    Returns dict with 'corrected_text' and 'changes_made', or None if no API key.
    """
    if not ENABLE_CLAUDE_AI:
        logging.info("Claude AI validation is disabled (ENABLE_CLAUDE_AI=False)")
        return None

    client = get_ai_client()
    if client is None:
        return None

    # Data classification warning for large documents
    text_len = len(document_text)
    if text_len > 50000:
        logging.warning(
            f"Large document ({text_len} chars) being sent to external AI service. "
            "Ensure document classification permits external processing."
        )

    prompt = build_dynamic_prompt(ai_rules, document_text)

    logging.info(f"Calling Claude ({CLAUDE_MODEL} via {AI_PROVIDER}) with {text_len} chars, {len(ai_rules)} rules")

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=CLAUDE_MAX_TOKENS,
        temperature=CLAUDE_TEMPERATURE,
        messages=[{"role": "user", "content": prompt}]
    )

    # Track token usage for monitoring (SOC 2 CC7.2)
    usage = getattr(response, 'usage', None)
    if usage:
        logging.info(f"Claude tokens — input: {usage.input_tokens}, output: {usage.output_tokens}")

    response_text = response.content[0].text
    json_start = response_text.find('{')
    json_end = response_text.rfind('}') + 1

    if json_start >= 0 and json_end > json_start:
        json_text = response_text[json_start:json_end]
        try:
            result = json.loads(json_text, strict=False)
        except json.JSONDecodeError:
            json_text = json_text.replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')
            result = json.loads(json_text)

        return {
            'corrected_text': result.get('corrected_text', ''),
            'changes_made': result.get('changes_made', 0)
        }

    raise ValueError("Could not parse JSON from Claude's response")
