"""Regenerate the Mace Style Validator diagrams via the imggen API (gemini-3-pro-image-preview).

Calls the imggen backend directly on the Mac Mini (bypasses Cloudflare Access).
Each run is non-deterministic — review the output. Saves diagram-*.jpg into this folder.
"""
import os, json, base64, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ENDPOINT = "http://192.168.5.49:3097/generate"   # Mac Mini imggen backend (LAN)

DIAGRAMS = {
    "diagram-architecture.jpg": """A clean, professional corporate software architecture diagram on a pure white background. Flat infographic style, crisp legible sans-serif labels, generous spacing, executive-grade, minimal. Brand colours: muted warm grey (#797574) and a deep green (#2E7D52) accent.

Layout: two large rounded-rectangle CONTAINER boxes side by side.

LEFT container, green outline, top label "MACE 365 TENANT". Inside it three stacked rounded boxes joined by small downward arrows:
1) green box: "Mace SharePoint" / "documents, rules, results, reports"
2) grey box: "Power Automate flow" / "runs as a Mace user, all writes"
3) dark grey box: "Entra app registration" / "Sites.Selected, one site, revocable"

RIGHT container, dark grey outline, top label "EXTERNAL AZURE SUBSCRIPTION". Inside it one dark grey box: "Azure Function" / "validation engine, stateless, stores nothing".

Between the two containers: two short horizontal arrows, top one pointing right, bottom one pointing left.

Bottom: a full-width light grey banner with green outline reading "Data never leaves the Mace tenant".

Wide 16:9 aspect ratio. Accurate, readable text only.""",

    "diagram-process.jpg": """A clean professional process-flow infographic on a pure white background. Flat style, crisp legible sans-serif labels, generous spacing, executive-grade. Brand colours: muted warm grey (#797574) and deep green (#2E7D52).

Six rounded rectangular step boxes connected by arrows. Top row, left to right: step 1, step 2, step 3 (grey boxes). Bottom row, right to left: step 4, step 5, step 6 (green boxes). A downward arrow connects step 3 to step 4. Arrows flow 1 to 2 to 3 to 4 to 5 to 6.

Each box has a bold number+title and a small caption:
1. "Validate Now" - user flags the document
2. "Flow triggers" - sends document to the engine
3. "Checked" - against the live style rules
4. "Auto-fixed" - issues corrected in a copy
5. "Report" - branded report generated and linked
6. "Recorded" - pass, review or fail logged

Bottom: a full-width light grey banner with green outline reading "The user only sets Validate Now - the rest is automatic".

Wide 16:9 aspect ratio. Accurate, readable text only.""",

    "diagram-pilot-prod.jpg": """A clean professional "before and after" comparison architecture diagram on a pure white background. Flat style, crisp legible sans-serif labels, executive-grade, minimal. Brand colours: muted warm grey (#797574) and deep green (#2E7D52).

Two side-by-side panels (rounded-rectangle containers).

LEFT panel, grey outline, title "TODAY - PILOT". Inside, stacked vertically: a green box "Mace 365 tenant" (data, SharePoint, flow), a downward arrow, then a dark grey box "Validation engine" (personal Azure subscription).

RIGHT panel, green outline, title "TARGET - PRODUCTION". Inside, stacked vertically: a green box "Mace 365 tenant" (data, SharePoint, flow), a downward arrow, then a green box "Validation engine" (Mace Azure subscription).

A large horizontal arrow points from the left panel to the right panel, labelled "Redeploy - same code".

Bottom: a full-width light grey banner with green outline reading "Same security model - only the hosting moves".

Wide 16:9 aspect ratio. Accurate, readable text only.""",
}


def generate(name, prompt):
    data = json.dumps({"prompt": prompt}).encode()
    req = urllib.request.Request(ENDPOINT, data=data, headers={"Content-Type": "application/json"})
    resp = json.load(urllib.request.urlopen(req, timeout=280))
    open(os.path.join(HERE, name), "wb").write(base64.b64decode(resp["image"]))
    print("wrote", name, "  source:", resp.get("url", "?"))


if __name__ == "__main__":
    for name, prompt in DIAGRAMS.items():
        generate(name, prompt)
