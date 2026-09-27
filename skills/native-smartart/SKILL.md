---
name: native-smartart
description: Create native, editable Microsoft Office SmartArt through NativeDiagram when the user requests SmartArt or editable process, hierarchy, cycle, or related structured diagrams for Word or PowerPoint. Generate standalone DOCX/PPTX diagrams or compose them into complete reports and slide decks using the bundled local insertion script and the host document workflow. Do not trigger for ordinary document writing or when the user specifies Mermaid, SVG, or another format.
---

# Native SmartArt

Use NativeDiagram's connected MCP tools to create the diagram portion of the user's task. Follow the user's chosen content, layout and delivery format. The surrounding document remains the responsibility of the host's document workflow.

## Scope and prerequisites

- The MCP creates a standalone DOCX and a single-slide PPTX with native SmartArt, plus a PNG preview. For complete reports or decks, read [composition guidance](references/composition.md) and use the bundled local insertion script. Host document tools own the surrounding content and layout.
- Use the tools exposed by the NativeDiagram connection, `list_templates` and `create_smartart`; their client-visible names may be namespaced. Do not invent tool names or assume another document skill is installed.
- If the connection is unavailable, explain that NativeDiagram must be connected through the client's normal OAuth flow. Never request passwords or API keys in chat.
- Each successful creation spends one credit and creates a new diagram. Discovering templates does not spend credits. Do not generate unsolicited variants to explore layouts.

## Preserve the content

Identify the requested relationship and prepare short labels with unique node keys. Follow the live tool schema and server instructions for roles and parent links. Preserve step order, hierarchy, names and significant distinctions.

Shorten wording only when the meaning survives. Keep explanatory detail in the surrounding prose when appropriate. Never drop steps, invent nodes or flatten a meaningful hierarchy just to fit a template. SmartArt's fixed structures cannot represent every graph; explain an unsupported relationship and let the user's intended output guide the next step.

## Choose a layout

When the template and its constraints are not already known, call `list_templates` for the intended structure. Supplying `nodes` filters the candidates by both relationships and text capacity. Omit `nodes` when first inspecting the supported shapes. The live response is authoritative; do not keep a template catalog in this skill.

Choose among fitting candidates using:

- Content relationships and slot limits, before visual preference.
- `layout_family`, dimensions and the intended location: a Word text column and a landscape slide have different proportions. Catalog dimensions describe the source diagram; PPTX output is scaled to fit its slide.
- Palette and layout consistency within the current document.

A thumbnail URL does not mean you have seen the image. Inspect a promising thumbnail with an available image-capable tool when visual comparison would change the choice; otherwise rely on metadata and do not claim visual inspection.

For multiple diagrams, remember the selected template IDs and layout families within this task. Use `exclude_template_ids` to find alternatives when variety helps. Exclusions are strict: if none remain, relax them deliberately when reuse is appropriate. Different IDs may share a layout family; a color change alone is not a new layout.

Keep the same layout for before/after comparisons and text-only revisions unless the user requests otherwise. Reuse a suitable template when alternatives do not fit. Do not rotate templates based on unrelated user history or select randomly.

## Create and deliver

Call `create_smartart` with the selected `template_id` and prepared nodes. If the user simply wants the default fitting layout, omitting `template_id` is supported. Use the returned actual template ID for later references.

The result includes a preview and structured file details: `diagram_id`, `template_id`, `structure`, `files`, `expires_at`, and `credits_charged`. Use the returned URLs exactly. Do not fabricate local paths or imply a link has already been downloaded.

Provide the preview and the requested editable file format; mention the other format briefly when useful. The preview is approximate and does not itself prove Office editability. If files can be inspected, verify the native SmartArt parts and inspect the rendered result using the host's available document tools. State the checks actually performed; do not claim testing in Word or PowerPoint without doing it.

For a complete report, presentation, or insertion request, first read [composition guidance](references/composition.md). Discover the host's document/presentation skill and runtime, create the full draft with an explicit marker, generate and download the diagram, then use `scripts/insert_smartart.py` to compose a new file. Render and inspect the complete result. Do not replace native SmartArt with PNG or improvise a merger. If the target or runtime is outside the supported contract, explain that before generating solely for insertion. The target theme is retained; source colors may change.

## Errors and revisions

- For input or capacity errors, consult the returned constraints and correct the specific issue without losing meaning. If there is no faithful representation, explain the limitation instead of cycling through retries.
- For quota exhaustion, explain the reported entitlement limit and stop. Do not advertise plans, initiate checkout, or add purchase links to tool responses.
- For a concurrency limit, wait for an outstanding request to complete; do not create a retry loop.
- A reported preview conversion failure means no diagram was saved or credit charged. A later retry may be appropriate if the service recovers.
- An ambiguous timeout or disconnected response may follow a completed, charged generation. There is no idempotency key or artifact-recovery tool in this version: do not automatically submit the same creation again. Explain the uncertainty and point to the user's NativeDiagram dashboard to check the result.
- A requested content or layout revision is a new creation and consumes another credit. Retain the chosen layout for text-only revisions when it still fits. Do not regenerate an unchanged result just to refresh its link; the dashboard can issue fresh download links.
