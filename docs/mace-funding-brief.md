# MaceStyle — Bringing It In-House

### An AI Service Lead brief for Tobias

**Prepared by Stephen Cummins · AI Service Lead, Mace Digital · 2 July 2026**

---

## The one-line version

MaceStyle — the automated Style Guide validator already piloting on the Mace Way Control Centre — can run entirely on **Mace's own Azure billing and governance**, at **pennies per document**, with **zero code changes**. I've already proven the switch end-to-end. It needs a Mace Azure subscription to own it and a data-governance nod. Then it's roughly half a day to cut over.

---

## What we have today

MaceStyle reads documents in the Mace Way Control Centre and checks them against the Writing Style Guide — British spelling, contractions, symbols, number and font standards — flagging issues and proposing corrections. It is **live in pilot**, validated by testers (Natasha, Jade) in June. The "intelligent" rules are powered by Claude (Anthropic's AI), which handles the judgement calls a simple find-and-replace can't.

**The catch:** the AI currently runs on *my personal* Anthropic account and Azure subscription. That's fine to prove the concept — it's not how Mace should run a service it depends on. To make this a Mace service, the billing and governance need to sit inside Mace.

## The proposal: run it on Mace's Azure, via Microsoft Foundry

Microsoft Foundry now offers Claude directly inside Azure. That changes the commercial picture entirely:

| | Today (my accounts) | Proposed (Mace) |
|---|---|---|
| **Who pays** | Me, personally | **Mace's existing Azure invoice** |
| **How** | Separate Anthropic account | Azure Marketplace — a line on the bill we already pay |
| **Governance** | None to speak of | **Entra ID, RBAC, data zone, spend caps** — Azure-native |
| **Procurement** | — | **None.** No new vendor, no new contract |

The key point for Mace: **no new supplier to onboard and no separate AI contract to negotiate.** Claude usage simply appears as consumption on the Azure agreement Mace already has. Microsoft handles billing and governance; Anthropic operates the model.

## What it costs

Consumption-based, and genuinely cheap at this scale. A document validation is a few thousand tokens of AI — **fractions of a penny to a couple of pence per document**. Even at hundreds of documents a day, this is a rounding error against the value of consistent, on-brand documentation across the Control Centre. **No upfront commitment; Mace pays only for what it uses.**

## What I need from you

Three decisions — none of them heavy:

1. **A home for it** — nominate a Mace Azure subscription and cost centre to own the service. Mace already runs on Azure (the `mace365` tenant, Power Platform), so nothing new is stood up.
2. **A governance nod** — document text is processed by a US-hosted model. Foundry's US data zone and Azure governance are built for exactly this; it needs Sapna's sign-off, which I'll tee up.
3. **Half a day of my time** to cut over — I've already rehearsed the entire handover in a test environment, so this is execution, not discovery.

## Why this matters beyond MaceStyle

MaceStyle is the **proof, not the point.** It shows a repeatable pattern: identify a real Mace problem, stand up an AI service to solve it, and hand it over running on Mace's own billing and governance — safely, cheaply, and without a procurement cycle.

That pattern is the AI Service Lead role in practice. Give me the mandate and MaceStyle becomes the first of several — each one an internal AI service that pays for itself in hours saved and lands inside Mace's existing Azure controls from day one.

---

*Detail for whoever provisions it: see the companion **Foundry Handover Runbook** (`docs/foundry-handover-runbook.md`).*
