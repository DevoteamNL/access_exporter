---
name: token-saver
description: Route work to the cheapest model that can do it well. USE FOR - deciding whether a step should be handed to the Haiku helpers (Scout for read-only lookups, Runner for lint/type-check/test/build runs) or a lower reasoning effort, and when to escalate back to a stronger model. Apply in any coding session, planning, review or research task to cut token use without lowering quality.
---

# Token saver: model routing

Goal: same quality, fewer expensive tokens. Strong models (Opus) do judgement; cheap models (Haiku) do mechanical
work; Sonnet does most implementation. Every hand-off must have a clear question and a short, structured answer.

## Model tiers

| Tier | Model | Use for |
|------|-------|---------|
| Judgement | Claude Opus 5.5 | Planning, architecture, code review, live-data/security risk, debugging after 2 failed attempts |
| Build | Claude Sonnet 5.5 | Implementing an approved plan, research with sources, docs writing (NL/EN) |
| Mechanical | Claude Haiku 5.5 | Lookups, running checks, summarising output, renames, boilerplate from a clear pattern |

## Hand to **Scout** (Haiku, read-only) when the question

- has one factual answer in the workspace: "where is X defined", "who calls Y", "list all uses of Z";
- is a lookup in a spec or config file (OpenAPI operation, `requirements.txt`, `build.ps1`);
- asks for a summary of one file you do not need to read in full yourself.

Run independent Scout questions in parallel. If Scout answers `NOT FOUND`, `NEEDS JUDGEMENT` or confidence `low`,
do the work yourself or ask the stronger researcher — do not retry Scout with the same question.

## Hand to **Runner** (Haiku) when you need to

- run ruff, pyright, pytest, a build or an install and only need pass/fail plus failing lines;
- re-run checks after a fix.

Read only Runner's Failures section; fix the code yourself (Runner never fixes).

## Use a lower reasoning effort (when your tool supports `reasoning_effort`) for

- formatting, renaming, moving code, adding a changelog line, filling a known template;
- translating already-final English UI text into Dutch documentation (or the reverse) — then let the docs writer
  check it.

## Keep on the strong model — never delegate

- decisions that touch live data writes, secrets/keyring, saved profiles or release/tag/push steps;
- the final review verdict and the plan itself;
- anything where a wrong cheap answer costs more than the tokens saved.

## Escalation rule

If a Sonnet or Haiku agent fails the same step twice (tests still red, wrong answer, guessed API behaviour), stop
and escalate: hand the step with the evidence to the Opus agent (Planner/Reviewer), or tell the user to switch the
session model to Opus for that step.

## Model hints in plans

Planners tag each implementation step `[Sonnet]` (routine, follows an existing pattern) or `[Opus]` (concurrency,
live-data writes, security, design choices). Implementers run `[Sonnet]` steps themselves and announce `[Opus]`
steps so the user can switch the model for just that step.

## Cheapest token is the one not spent on rework

Unclear requirements are the biggest waste. Before planning non-trivial work, run the `grill-me` skill and capture
numbered acceptance criteria; implementers test each one and reviewers check each one.

## Also cheap by habit

- Ask for structured, short outputs (tables, `file:line`) instead of prose.
- Read files with line ranges; do not re-read files you already have.
- Pipe verbose commands through Runner instead of reading full logs.
- One focused question per sub-agent call; give it all context so it does not have to explore.
