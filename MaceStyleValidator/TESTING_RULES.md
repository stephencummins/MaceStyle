# Verifying that style rules actually apply

Three small tools answer the question "will the rules really fire on a real
document?". They run locally — no Azure deploy needed.

## 1. Does the validation engine work? (offline, no creds)

```bash
python3 test_rule_coverage.py
```

Builds one `.docx` containing a known violation for every implemented Word
check, runs the real validator, and asserts each one is caught. Exits non-zero
if any check stops firing — use it as a regression gate before deploying.

This proves the **code** works. It does **not** prove your SharePoint rules are
shaped to reach it — for that, see step 3.

## 2. Snapshot the real rules (needs SharePoint creds)

On a machine where the `SHAREPOINT_*` / `STYLE_RULES_*` env vars are set
(e.g. exported from `local.settings.json`):

```bash
python3 dump_style_rules.py            # writes rules_snapshot.json
```

## 3. Will the REAL rules fire? (offline, on the snapshot)

```bash
python3 rule_doctor.py rules_snapshot.json
```

Statically checks every rule for the silent failure modes that make a rule a
no-op even when it looks fine in SharePoint:

- `auto_fix` / `use_ai` stored as the **string** `"Yes"`/`"No"` instead of a
  boolean — any non-empty string is truthy in Python, so a `"No"` auto-fix
  still fires and a `"No"` UseAI still routes the rule to the AI path.
- `doc_type` that no validator matches (Word docs only see `Word` / `Both` /
  `All`) — silently dropped.
- `rule_type` the dispatcher doesn't recognise — silently dropped.
- `check_value` with no validator branch — silently skipped.

It prints how many rules will actually apply to Word documents, and lists the
ones to fix in the SharePoint **Style Rules** list.

Run `python3 rule_doctor.py` with no argument to see what a healthy/failing
result looks like against a built-in demo set.

## Likely cause of the early test-document failures

`test_rule_coverage.py` shows the engine catches all implemented checks, so a
real document that "passes" when it shouldn't almost always means the rule data
is the problem — most often `doc_type` not in `{Word, Both, All}`, or a
`check_value` that doesn't match a validator branch. Run the doctor on a live
snapshot to confirm.
