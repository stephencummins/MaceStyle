"""Centralised Claude AI client for style validation"""
import os
import json
import logging
from anthropic import Anthropic
from .config import (
    ENABLE_CLAUDE_AI, AI_PROVIDER, CLAUDE_MODEL, CLAUDE_MAX_TOKENS, CLAUDE_TEMPERATURE,
    AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, AZURE_OPENAI_API_VERSION,
    AZURE_OPENAI_MAX_COMPLETION_TOKENS, AZURE_OPENAI_REASONING_EFFORT,
)


def get_ai_client():
    """Return an AI client for the configured provider, or None if unconfigured.

    The two Anthropic providers (anthropic, foundry) expose the identical
    client.messages.create() surface; azure_openai exposes the OpenAI
    chat.completions surface instead. _generate() below hides that difference so
    the rest of this module stays provider-agnostic. Switching providers is an
    app-settings change only (AI_PROVIDER + the provider's creds + CLAUDE_MODEL) -
    no code change.
    """
    if AI_PROVIDER == "azure_openai":
        from openai import AzureOpenAI
        if not (AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY):
            logging.warning("AZURE_OPENAI_ENDPOINT/AZURE_OPENAI_API_KEY not set - skipping AI validation")
            return None
        return AzureOpenAI(
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            api_key=AZURE_OPENAI_API_KEY,
            api_version=AZURE_OPENAI_API_VERSION,
        )

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


def _generate(client, prompt):
    """Run one completion and return (response_text, input_tokens, output_tokens).

    Hides the Anthropic Messages API vs OpenAI Chat Completions difference so the
    caller only deals with text. CLAUDE_MODEL carries the model/deployment name
    for whichever provider is active.
    """
    if AI_PROVIDER == "azure_openai":
        # GPT-5 reasoning models: use max_completion_tokens (not max_tokens), leave
        # temperature at its default (custom values are rejected), and ask for JSON.
        response = client.chat.completions.create(
            model=CLAUDE_MODEL,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            max_completion_tokens=AZURE_OPENAI_MAX_COMPLETION_TOKENS,
            reasoning_effort=AZURE_OPENAI_REASONING_EFFORT,
        )
        usage = getattr(response, "usage", None)
        in_tok = getattr(usage, "prompt_tokens", None) if usage else None
        out_tok = getattr(usage, "completion_tokens", None) if usage else None
        return response.choices[0].message.content, in_tok, out_tok

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=CLAUDE_MAX_TOKENS,
        temperature=CLAUDE_TEMPERATURE,
        messages=[{"role": "user", "content": prompt}],
    )
    usage = getattr(response, "usage", None)
    in_tok = getattr(usage, "input_tokens", None) if usage else None
    out_tok = getattr(usage, "output_tokens", None) if usage else None
    return response.content[0].text, in_tok, out_tok


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

    logging.info(f"Calling AI ({CLAUDE_MODEL} via {AI_PROVIDER}) with {text_len} chars, {len(ai_rules)} rules")

    response_text, in_tok, out_tok = _generate(client, prompt)

    # Track token usage for monitoring (SOC 2 CC7.2)
    if in_tok is not None or out_tok is not None:
        logging.info(f"AI tokens — input: {in_tok}, output: {out_tok}")

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
