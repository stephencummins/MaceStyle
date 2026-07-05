# MaceStyle — Bringing It In-House

### An AI Service Lead brief for Tobias

**Prepared by Stephen Cummins · AI Service Lead, Mace Digital · 2 July 2026 (updated 5 July 2026)**

---

## The one-line version

MaceStyle — the automated Style Guide validator already piloting on the Mace Way Control Centre — now runs its AI on **GPT via Azure OpenAI**: Microsoft-native, on Azure billing and governance, at **pennies per document**. I've already made the switch and proven it end-to-end on a live pilot. To make it a Mace service it needs a Mace Azure subscription to own it and a data-governance nod. Then it's roughly half a day to cut over.

---

## What we have today

MaceStyle reads documents in the Mace Way Control Centre and checks them against the Writing Style Guide — British spelling, contractions, symbols, number and font standards — flagging issues and proposing corrections. It is **live in pilot**, validated by testers (Natasha, Jade) in June. The "intelligent" rules — the judgement calls a simple find-and-replace can't make — are powered by AI.

**What changed this month:** I re-architected MaceStyle so the AI backend is *swappable* via a single setting, and switched the live pilot onto **GPT-5 (Azure OpenAI)** — Microsoft's own AI service, running inside Azure. It's proven working end-to-end. The same code can run Claude instead with no rewrite, so we are not locked to one model.

**The catch:** it currently runs on *my personal* Azure subscription. That's fine to prove the concept — it's not how Mace should run a service it depends on. To make this a Mace service, the billing and governance need to sit inside Mace.

## The proposal: run it on Mace's Azure — Microsoft-native

Mace's instinct is to stay on Microsoft wherever possible. This fits that exactly. The AI now runs on **Azure OpenAI** — the GPT models (the ChatGPT family) delivered as a first-party Microsoft Azure service. That means:

| | Today (my Azure) | Proposed (Mace's Azure) |
|---|---|---|
| **Who pays** | Me, personally | **Mace's existing Azure invoice** |
| **What it is** | GPT via Azure OpenAI | Same — a native Azure service Mace already has access to |
| **Governance** | Mine | **Entra ID, RBAC, data zone, spend caps** — Azure-native |
| **Procurement** | — | **None.** No new vendor, no new contract, no third-party sign-up |

The key point for Mace: **there is no new supplier and no AI contract to negotiate.** Azure OpenAI is Microsoft, on the Azure agreement Mace already runs. Standing it up in a Mace subscription is deploying a model and pointing the app at it — no marketplace step, no external account.

**Optional — model choice stays open.** Because the backend is swappable, Mace can run **Claude** instead (via Microsoft Foundry, also inside Azure) for any content that benefits from its slightly more careful editing. That's a one-setting change, not a rebuild. My recommendation is to **default to GPT/Azure OpenAI** — it's the easiest for Mace to provision and the cheaper of the two — and keep Claude available as an option.

## What it costs

Consumption-based and genuinely cheap. I measured both models on real documents through the live pipeline:

| Document size | GPT-5-mini (recommended) | Claude (alternative) |
|---|---|---|
| ~1 page | **£0.003** | £0.004 |
| ~3 pages | **£0.004** | £0.008 |
| ~8 pages | **£0.008** | £0.015 |

*(US-dollar list prices, converted approximately; Mace's enterprise Azure agreement would likely discount further.)*

GPT is **~35–50% cheaper per document** than Claude, and the gap widens with document length. But in absolute terms both are a **rounding error**: realistic Control Centre volume is **a few pounds a month**. There is **no upfront commitment — Mace pays only for what it uses.** Cost is not the deciding factor here; ease of ownership is, and Azure OpenAI wins on both.

## What I need from you

Three decisions — none of them heavy:

1. **A home for it** — nominate a Mace Azure subscription and cost centre to own the service. Mace already runs on Azure (the `mace365` tenant, Power Platform), so nothing new is stood up.
2. **A governance nod** — document text is processed by a US-hosted model. Azure OpenAI's data zone and Azure-native governance are built for exactly this; it needs Sapna's sign-off, which I'll tee up.
3. **Half a day of my time** to cut over — I've already built and proven the whole path on a live pilot, so this is execution, not discovery.

## Why this matters beyond MaceStyle

MaceStyle is the **proof, not the point.** It shows a repeatable pattern: identify a real Mace problem, stand up an AI service to solve it, and hand it over running on Mace's own Microsoft/Azure billing and governance — safely, cheaply, and without a procurement cycle.

That pattern is the AI Service Lead role in practice. Give me the mandate and MaceStyle becomes the first of several — each one an internal AI service that pays for itself in hours saved and lands inside Mace's existing Azure controls from day one.

---

*Detail for whoever provisions it: see the companion **Foundry Handover Runbook** (`docs/foundry-handover-runbook.md`), which covers both the Azure OpenAI (GPT) and Foundry (Claude) provider setups.*
