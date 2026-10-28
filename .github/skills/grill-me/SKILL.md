---
name: grill-me
description: Grill the user relentlessly about a plan, feature request, issue or design until requirements are unambiguous, before any plan or code is written. USE FOR - "grill me", "stress-test this idea", unclear feature requests or issues, and as the first step of planning a feature or release to prevent rework. DO NOT USE FOR - small, fully specified changes.
---

<!--
Adapted from "grilling" / "grill-me" by Matt Pocock — https://github.com/mattpocock/skills (MIT License,
Copyright (c) Matt Pocock). Changes: cheap fact-finding via the Scout agent, requirement checklist, output summary.
-->

# Grill me

Interview the user relentlessly until you reach a shared understanding. Map this as a **design tree**: every
decision branches into the decisions that hang off it.

## Rounds

Work the tree in **rounds**. The **frontier** is every decision whose prerequisites are already settled: the
questions you can ask _now_ without guessing at answers you haven't heard yet. Ask the whole frontier in one round:
number each question and give your recommended answer. Then wait for the user's answers before the next round.

Format a round like so:

```
❓ **Q1** - **<question title>**: <question body, might be multiple paragraphs, including multiple choices>

➡️ <your recommended answer>

---

❓ **Q2** - **<question title>**: <question body>

➡️ <your recommended answer>
```

Word each question so "yes" accepts your recommended answer. Keep a round to at most 7 questions; ask the most
consequential ones first.

Each round the user answers reshapes the tree: settled decisions push the frontier outward and unblock questions
that depended on them. Recompute the frontier and ask the next round. A question whose answer depends on another
question still open in this round belongs to a _later_ round, not this one.

## Facts are your job, decisions are the user's

Finding _facts_ is your job, never the user's. When a frontier question needs a fact from the code, docs or tools,
dispatch the cheap **Scout** agent (Haiku, read-only) to find it — or a stronger researcher when the fact needs
judgement (for example API behaviour). Don't ask the user for anything you could look up yourself. Don't block on
it: only the questions downstream of a running lookup wait; ask the rest of the frontier now.

The _decisions_ are the user's: put each to them and wait.

## Branches to cover (check each before you stop)

- Goal and user: who needs this, and what problem disappears when it is done?
- Scope: what is explicitly in, and explicitly out (for this release)?
- Behaviour: happy path, empty/large data, errors, cancel, permissions.
- Exact UI text and labels (all languages the product uses).
- Data safety: does it change live/production data? What preview or confirmation is required?
- Secrets and stored settings: any impact on credentials, saved profiles or config files?
- Compatibility: existing files, settings, APIs and versions that must keep working.
- Acceptance criteria: how will the user verify it is done? Which tests prove it?
- Non-functional: performance, packaging/dependencies, platforms.

## Done

The session is done when the frontier is empty: every branch of the design tree visited, nothing left silently
assumed. Then write a short **Requirements summary** (decisions, out-of-scope, acceptance criteria, remaining
risks) and ask the user to confirm it. Do not plan or implement until the user confirms you have reached a shared
understanding. A planner should copy the confirmed summary into the plan so the developer never has to guess.

In a non-interactive context (e.g. a cloud agent working from an issue), do not wait: put the rounds as open
questions with recommended defaults in the pull request description instead.
