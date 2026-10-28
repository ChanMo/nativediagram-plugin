---
name: native-smartart
description: Create native, editable Microsoft Office SmartArt through NativeDiagram when the user requests SmartArt or editable process, hierarchy, cycle, or related structured diagrams for Word or PowerPoint. Generate standalone DOCX/PPTX diagrams or compose them into complete reports and slide decks using the bundled local insertion script and the host document workflow. Do not trigger for ordinary document writing or when the user specifies Mermaid, SVG, or another format.
---

# Native SmartArt

Use NativeDiagram's connected MCP tools to create the diagram portion of the user's task. Follow the user's chosen content, layout and delivery format. The surrounding document remains the responsibility of the host's document workflow.

## Scope and prerequisites

- The MCP creates a standalone DOCX and a single-slide PPTX with native SmartArt, plus a PNG preview. For complete reports or decks, read [composition guidance](references/composition.md) and use the bundled local insertion script. Host document tools own the surrounding content and layout.
- Use the tools exposed by the NativeDiagram connection: `list_templates`, `get_template_previews`, `create_smartart`, and `get_smartart_result`. Their client-visible names may be namespaced. Do not invent tool names or assume another document skill is installed.
- If the connection is unavailable, explain that NativeDiagram must be connected through the client's normal OAuth flow. Never request passwords or API keys in chat.
- Each new successful creation spends one credit. Recovering or replaying the same keyed result is free. Discovering templates and viewing sample previews do not spend credits. Do not generate unsolicited variants to explore layouts.

## Preserve the content

Identify the requested relationship and prepare short labels with unique node keys. Follow the live tool schema and server instructions for roles and parent links. Preserve step order, hierarchy, names and significant distinctions. Preserve the user’s requested content language, including mixed-language labels; do not translate labels to fit a template or infer their language from this skill’s English text.

Shorten wording only when the meaning survives. Keep explanatory detail in the surrounding prose when appropriate. Never drop steps, invent nodes or flatten a meaningful hierarchy just to fit a template. SmartArt's fixed structures cannot represent every graph; explain an unsupported relationship and let the user's intended output guide the next step.

## Choose a layout

When the template and its constraints are not already known, call `list_templates` for the intended structure. Supplying `nodes` filters the candidates by both relationships and text capacity. Omit `nodes` when first inspecting the supported shapes. The live response is authoritative; do not keep a template catalog in this skill.

For `relationship`, explicitly pass `relation="overlap"` for overlapping sets or `relation="equation"` for A + B = C; in an equation the center is the result and factors are the inputs. Generic cooperation fits neutral radial associations. For `radial`, omitted relation means `association`; choose `converging` or `diverging` only when the user's content requires that arrow direction. Inspect `relation`, `directed`, `cyclic`, `best_for` and `avoid_for`, and pass the same relation to `create_smartart`, including when selecting an explicit template ID. Schematic timelines do not encode proportional dates or durations, matrices do not encode numeric axes, and conditional links, merges and swimlanes are unsupported.

Choose among fitting candidates using:

- Content relationships and slot limits, before visual preference.
- With `nodes`, inspect `preview_text`: minimum/maximum font points from the filled standalone DOCX drawing cache, and the count of shapes whose size is inherited/unknown. Prefer a roomier layout when text would shrink excessively. Null sizes do not establish readability; review the image. These sizes precede PPTX scaling and final insertion resizing, and do not measure the resolved installed font.
- If the intended space is known, pass `placement` with full nodes: `{format: "docx"|"pptx", width_cm, height_cm?, min_font_pt?}`. PPTX requires height; Word uses the exact width and reports optional height overflow. Compare `placement_fit` estimated dimensions/font sizes with the user’s context, and reuse those dimensions for insertion. Null `meets_min_font` is unknown or unrequested. The query keeps unsuitable/unknown estimates visible rather than silently filtering them. Estimates do not resolve host fonts/themes or Office reflow; render the composed result. Do not impose a universal minimum font size.
- `layout_family`, dimensions and the intended location: a Word text column and a landscape slide have different proportions. Catalog dimensions describe the source diagram; PPTX output is scaled to fit its slide.
- Palette and layout consistency within the current document. Prefer the existing document's theme and the user's explicit brand/print/background requirements. Keep equal peers visually equal; use color emphasis only when the content calls for it.

English is the primary product validation language, with Chinese and Japanese also supported as target content languages. Slot limits use estimated character widths, not measured typography. For English, inspect long words, phrase wrapping and readable font size; for Chinese/Japanese, inspect full-width text, kana, punctuation and missing glyphs. Check Latin and East Asian font/theme settings separately in the actual file. A short Chinese sample does not establish English or Japanese fit. Preserve exact requested labels and prefer a roomier fitting layout over excessive font shrinking.

Native `style` metadata includes theme colors, cached color references with their transforms, applied effects and native style IDs. `palette_name` is a descriptive label, not proof of the visible colors. Theme colors are not necessarily used by every shape; `base_hex` precedes any listed color transforms. Do not treat an internal style ID as a known aesthetic. Colors/effects are fixed by the template; this version has no independent recoloring or restyling parameters.

When appearance would change the choice, call `get_template_previews` with up to three fitting IDs. Each caption identifies the following image. These cached thumbnails contain sample labels, not the user's content; their original renderer/DPI may be unknown. Compare proportions, color/contrast and effects alongside the native metadata. A thumbnail URL alone does not mean you have seen the image. If the host/model cannot inspect images, use metadata and state that visual inspection was not performed.

For multiple diagrams, remember the selected template IDs and layout families within this task. Use `exclude_template_ids` to find alternatives when variety helps. Exclusions are strict: if none remain, relax them deliberately when reuse is appropriate. Different IDs may share a layout family; a color change alone is not a new layout.

Keep the same layout for before/after comparisons and text-only revisions unless the user requests otherwise. Reuse a suitable template when alternatives do not fit. Do not rotate templates based on unrelated user history or select randomly.

## Create and deliver

Before the first creation, generate and retain a unique UUID `request_id` together with the exact input. Call `create_smartart` with this ID, the selected `template_id` and prepared nodes. If the user simply wants the default fitting layout, omitting `template_id` is supported. Use the returned actual template ID for later references.

The result includes a filled preview and structured file details: `diagram_id`, `template_id`, `structure`, `files`, `preview`, `expires_at`, `request_id`, `reused`, and `credits_charged`. Use the returned URLs exactly. Do not fabricate local paths or imply a link has already been downloaded.

Before delivery, inspect the filled image when image capabilities are available. Compare it against ALL original nodes: text completeness, duplicates, order, parent links and arrow meaning; then check wrapping, occlusion, contrast and visual weight. Cross-check ambiguous text with the file rather than guessing from a small image. Do not shorten exact labels or delete nodes to make the picture look better.

The generated `preview` identifies a cropped standalone DOCX rendered by LibreOffice. It is not a PPTX preview, a final-page check or proof of native Office editability. A cropped diagram also cannot establish readability at its final insertion size. For a PPTX task, inspect the downloaded PPTX using host presentation tools; for a composed document, render its actual pages/slides after insertion. Source-theme colors can change under the target theme. State the checks actually performed; do not claim Word/PowerPoint testing without doing it.

If the image reveals a problem, identify whether it concerns content, layout, document placement or renderer differences. Query suitable alternatives for free before changing the template. Re-render the same downloaded file to investigate a suspected renderer issue; reuse it for surrounding-document layout changes. Respect requested revisions and the user's budget; do not make charged attempts solely to explore appearance. Provide the preview and requested editable format, and mention the other format when useful.

For a complete report, presentation, or insertion request, first read [composition guidance](references/composition.md). Discover the host's document/presentation skill and runtime, create the full draft with an explicit marker, generate and download the diagram, then use `scripts/insert_smartart.py` to compose a new file. Render and inspect the complete result. Do not replace native SmartArt with PNG or improvise a merger. If the target or runtime is outside the supported contract, explain that before generating solely for insertion. The target theme is retained; source colors may change.

## Errors and revisions

- For input or capacity errors, consult the returned constraints and correct the specific issue without losing meaning. If there is no faithful representation, explain the limitation instead of cycling through retries.
- For quota exhaustion, explain the reported entitlement limit and stop. Do not advertise plans, initiate checkout, or add purchase links to tool responses.
- For a concurrency limit, wait for an outstanding request to complete; do not create a retry loop.
- A reported preview conversion failure means no diagram was saved or credit charged. A later retry may be appropriate if the service recovers.
- After an ambiguous timeout or disconnected response, call `get_smartart_result` with the retained `request_id`. `ready` returns the original files/preview and fresh links for free; `pending` needs time. For `retryable`, resubmit the exact original input with the SAME ID. For `not_found`, a delayed call may still arrive; the same keyed input is safe, but do not switch IDs. Changed input needs a new ID. If the live connection lacks keyed creation/recovery, or the original call omitted an ID and no diagram ID was returned, check the dashboard rather than blindly repeat a charge. The live deployed schema is authoritative.
- A requested content or layout revision is a new creation and consumes another credit. Retain the chosen layout for text-only revisions when it still fits. Do not regenerate an unchanged result just to refresh its link; call `get_smartart_result` with the known request or diagram ID for free fresh links; the dashboard also supports link refresh.
