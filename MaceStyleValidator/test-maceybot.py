#!/usr/bin/env python3
"""Local test script for MaceyBot conversation flow.

Simulates a Teams conversation by calling Claude API directly.
Requires ANTHROPIC_API_KEY environment variable.

Usage:
    python3 test-maceybot.py
"""
import json
import os
import sys

from anthropic import Anthropic

# Import the prompts and tools from the bot
sys.path.insert(0, os.path.dirname(__file__))
from MaceyBot.claude_client import SYSTEM_PROMPT, TOOLS

MODEL = os.environ.get("MACEY_MODEL", "claude-sonnet-4-20250514")


def main():
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("Error: Set ANTHROPIC_API_KEY environment variable")
        sys.exit(1)

    client = Anthropic(api_key=api_key)
    messages = []

    print("=" * 60)
    print("MaceyBot Local Test — type 'quit' to exit")
    print("=" * 60)
    print()

    while True:
        user_input = input("You: ").strip()
        if not user_input:
            continue
        if user_input.lower() in ("quit", "exit", "q"):
            break

        messages.append({"role": "user", "content": user_input})

        response = client.messages.create(
            model=MODEL,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages,
        )

        print(f"  [stop_reason={response.stop_reason}, "
              f"tokens={response.usage.input_tokens}in/{response.usage.output_tokens}out]")

        if response.stop_reason == "tool_use":
            # Print any text before the tool call
            for block in response.content:
                if block.type == "text" and block.text:
                    print(f"Macey: {block.text}")
                elif block.type == "tool_use":
                    print(f"\n  >>> TOOL CALL: {block.name}")
                    print(f"  >>> Params: {json.dumps(block.input, indent=2)}")

                    # Simulate successful submission
                    tool_result = json.dumps({
                        "status": "success",
                        "message": "Site creation request submitted successfully (ID: TEST-001). "
                                   "The approval workflow has been triggered.",
                        "item_id": "TEST-001",
                    })
                    print(f"  >>> Result: {tool_result}")

                    # Add tool interaction to history
                    messages.append({"role": "assistant", "content": response.content})
                    messages.append({
                        "role": "user",
                        "content": [{
                            "type": "tool_result",
                            "tool_use_id": block.id,
                            "content": tool_result,
                        }],
                    })

                    # Get final response
                    final = client.messages.create(
                        model=MODEL,
                        max_tokens=1024,
                        system=SYSTEM_PROMPT,
                        tools=TOOLS,
                        messages=messages,
                    )

                    final_text = " ".join(
                        b.text for b in final.content if hasattr(b, "text") and b.text
                    )
                    print(f"\nMacey: {final_text}")
                    messages.append({"role": "assistant", "content": final_text})
        else:
            text = " ".join(
                b.text for b in response.content if hasattr(b, "text") and b.text
            )
            print(f"Macey: {text}")
            messages.append({"role": "assistant", "content": text})

        print()


if __name__ == "__main__":
    main()
