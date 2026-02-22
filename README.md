# PPT Automation

AI-assisted PowerPoint recommendations and safe slide updates.

## Vision

Read what is happening on a slide, infer user intent, recommend next actions, and apply accepted actions without breaking the slide's current design.

This is domain-agnostic. It is not limited to states/capitals or any single dataset.

## Architecture

### High-level components

1. PowerPoint Add-in (Office.js task pane)
- Reads active slide and selected objects.
- Sends normalized slide context + user prompt to backend.
- Shows ranked recommendations.
- Sends accepted recommendation to "plan" API.
- Applies returned execution plan transactionally.

2. Backend API (Node.js)
- Intent and context understanding.
- Recommendation generation (LLM + deterministic shaping).
- Execution plan generation (LLM + rule-based constraints).
- Validation and policy checks.

3. Local LLM Runtime (Ollama)
- Default model: `qwen2.5:7b-instruct`.
- Structured JSON outputs for recommendations and plans.

4. Rendering/Apply Engine (in add-in runtime)
- Uses theme tokens and layout constraints.
- Avoids overlap and preserves existing design language.
- Supports preview, apply, undo.

### Data flow

1. Add-in captures `SlideContext`.
2. Client sends `SlideContext + userPrompt` to `POST /api/recommendations`.
3. Backend returns ranked recommendation cards.
4. User accepts one recommendation.
5. Client sends selection to `POST /api/plans`.
6. Backend returns deterministic `ExecutionPlan`.
7. Add-in applies plan with transactional semantics.

### Core contracts

1. `SlideContext`
- `slide.id`, `slide.index`, `slide.size`.
- `selection.shapeIds`.
- `objects[]`: type, geometry, text/table/chart metadata, style hints.
- `themeHints`.

2. `Recommendation`
- `id`, `title`, `description`, `outputType`, `confidence`, `applyHints`.

3. `ExecutionPlan`
- `operations[]` where each operation is one of: `insert`, `update`, `transform`, `delete`.
- `target`, `anchor`, `styleBindings`, `constraints`.
- `warnings[]`, `requiresConfirmation`.

## Design Plan

### Design principles

1. AI proposes, renderer applies.
- LLM never mutates PowerPoint directly.
- LLM output is constrained JSON only.

2. Preserve design first.
- Reuse placeholders and existing styles.
- Bind to slide theme tokens.
- Enforce no-overlap and no-cutoff rules.

3. Deterministic and reversible.
- Validate every plan before apply.
- Transactional apply and single-step undo.

4. Confidence-aware UX.
- High confidence: direct preview and apply.
- Medium confidence: show alternatives.
- Low confidence: ask clarification.

### Recommendation taxonomy (domain-agnostic)

- Summarize or expand content.
- Convert bullets to table.
- Convert table to chart.
- Create comparison layout.
- Create process or timeline structure.
- Extract action items / next steps.
- Improve visual hierarchy and readability.

## Implementation Plan

### Phase 1 (implemented in this repo now)

1. Backend scaffolding
- Express service with health endpoint.
- Recommendations endpoint using local Ollama.
- Execution plan endpoint using local Ollama.
- JSON parsing and fallback handling.

2. Prompting contracts
- System prompts that force strict JSON shape.
- Context-aware prompts with slide object summary.

3. Docs and local run
- Local setup for Ollama + model.
- API request/response examples.

### Phase 2 (partially implemented)

1. Strong validation layer
- Contract validation for `slideContext`, recommendations, and execution plans.
- Policy checks for destructive operations and free-region placement risk.
- Automatic fallback to safe plan/recommendations on invalid model output.

2. Office.js add-in context collector
- Extract text/table/chart/object metadata.
- Build normalized `SlideContext`.
- Initial scaffold added at `addin/src/slide-context-collector.js`.
- Implemented selected slide + selected shape extraction, object geometry, text/style hints, and table metadata with safe fallbacks.

### Phase 3

1. Add-in apply engine
- Preview diff and confirm.
- Theme-preserving render rules.
- Atomic apply + undo.
- Initial apply implemented: executes generated `insert/update/transform` text operations on the active slide with safe fallbacks.
- Placement upgrade implemented: prefers suitable placeholders (e.g., empty subtitle/content), estimates readable font size to fit bounds, and computes a non-overlapping free-region fallback.
- Insert behavior hardened: escaped newlines are normalized, unsafe insert targets (like authored titles) trigger alternative placeholder targeting before textbox fallback, and fallback textbox height adapts to content length.
- Title-protection policy added: body-like `transform/update` content (bullets/long text) is redirected away from title into body placeholder or a new textbox.
- Intent contract enforced in applier:
  - `transform/update` modifies only the explicit intended existing item (no heuristic reassignment).
  - `insert` adds new content without replacing authored slide content; it uses empty placeholders first, otherwise creates a new textbox in a safe region.
- Insert positioning upgrade:
  - New items are now placed using a scored layout strategy (title-aware, center-biased, overlap-penalized) rather than defaulting to the left corner.
- Confirm/reject workflow implemented: generated plans open a preview overlay in taskpane and only apply after explicit user confirmation.

2. Quality harness
- Regression slide fixtures.
- Plan quality scoring and telemetry.

## Local Setup

### Prerequisites

- Node.js 20+
- Ollama installed and running

### Install model

```powershell
ollama pull qwen2.5:7b-instruct
```

### Run backend

```powershell
cd server
npm install
npm run dev
```

Backend default URL: `http://localhost:4000`

### Run add-in web host

```powershell
cd addin
npm install
npm run dev
```

Add-in web host URL: `https://localhost:3100`

### Sideload into PowerPoint (desktop)

Keep backend and add-in web host running, then in a new terminal:

```powershell
cd addin
npm run sideload
```

The script uses `office-addin-debugging start manifest.xml desktop --app powerpoint`.

To stop sideload session:

```powershell
cd addin
npm run stop
```

## API (initial)

### `GET /health`
- Returns service status and model name.

### `POST /api/recommendations`
- Input: `{ userPrompt, slideContext }`
- Output: `{ recommendations: Recommendation[] }`

### `POST /api/plans`
- Input: `{ selectedRecommendation, userPrompt, slideContext }`
- Output: `{ plan: ExecutionPlan }`

### `POST /api/references`
- Input: `{ itemText, slideContext }`
- Output: `{ reference, alternatives }`

## Next repo tasks

1. Add JSON Schema validation and contract tests.
2. Expand add-in apply engine to execute plans on slide objects.
3. Add plan simulation and no-overlap checks before apply.
