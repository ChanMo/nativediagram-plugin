# Compose a complete report or presentation

Use the host's available document or presentation skill for writing and layout.
Do not assume a particular skill name, installation path, or authoring library.
The host must be able to download files, run Python 3 with `lxml`, and render the
finished document. Check these capabilities before spending a credit for an
insertion-only request. If they are unavailable, explain the limitation; do not
claim a complete document or silently substitute an image.

## Supported boundary

The bundled `scripts/insert_smartart.py` copies native diagram parts and their
relationships into an existing file. It runs locally, has no network access, and
never overwrites an existing output. It is not a general document merger.

- Input diagram: a standalone NativeDiagram DOCX or PPTX with exactly one SmartArt.
- Target must have the same format as the diagram.
- Word: a single-section document with a unique marker in a plain, top-level body
  paragraph. Use page breaks for a short multi-page report. Table cells, headers,
  multiple sections and multi-column layouts are outside this first contract.
- PowerPoint: a unique marker in a top-level text box with explicit position and
  dimensions. The marker is its text, not its shape name. No rotated/grouped boxes.
- The diagram adopts the target theme. Source colors/fonts may change; the target
  theme, styles, masters and unrelated content are not replaced. Exact preservation
  of the source theme is not supported.
- Multiple diagrams can be inserted sequentially at distinct markers. Use a new
  output file per operation. Duplicate or missing markers fail explicitly.
- Source diagram dependencies must be internal. Signed packages are unsupported.

## Workflow

1. Write the complete report or deck using the available host workflow. Reserve
   the intended space and add a unique marker such as `[[ND_PROCESS]]`.
2. Query and choose a NativeDiagram template that fits both the content and the
   reserved aspect ratio. Generate once; download the same-format diagram file.
3. Run the bundled insertion script using paths actually present in the workspace.
4. Render the final inserted file, inspect every page/slide and fix layout issues.
   If the host draft needs changes, update that draft and reinsert the downloaded
   diagram. Layout-only iterations must not create new charged generations.
5. Check the package still contains diagram data/layout/style/color parts with
   resolvable relationships, and no marker remains. Report which checks ran.
6. Deliver the complete DOCX/PPTX as the primary output. The standalone source is
   an intermediate, not a substitute for the user's requested report.

Do not pass the final inserted file through an unverified importer/exporter: it
may drop or flatten unsupported SmartArt. Renderers are for inspection, not for
replacing the editable file. Native editability in Microsoft Office requires an
actual Office check; XML presence and a rendered preview alone do not prove it.

## Word example

Create an A4 report with one section, a body marker paragraph and adequate vertical
space. Keep the marker paragraph separate from headings, captions and page breaks.
The optional width is in centimetres and must fit the page's content width. Aspect
ratio is retained; omitted width uses the smaller of source width and page width.
Page fitting is checked by rendering, not inferred by the insertion script.

```sh
python /path/to/native-smartart/scripts/insert_smartart.py \
  --target report-draft.docx --diagram diagram.docx \
  --marker '[[ND_PROCESS]]' --width-cm 14 --output report-final.docx
```

## PowerPoint example

Create the deck with the host's presentation tools. Add a plain text box containing
only `[[ND_PROCESS]]` where the diagram should sit. Its rectangle defines the
available space. The script keeps aspect ratio and centres the diagram in that box.

```sh
python /path/to/native-smartart/scripts/insert_smartart.py \
  --target deck-draft.pptx --diagram diagram.pptx \
  --marker '[[ND_PROCESS]]' --output deck-final.pptx
```

Replace the example interpreter and paths with the host's actual runtime and files.
If `lxml` is missing, use the host's normal dependency setup or report the missing
prerequisite. Do not attempt to reimplement the merger ad hoc.

## Revisions

Keep the clean draft and the downloaded diagram as working inputs. Text/layout
changes outside the diagram reuse the same downloaded source without another
credit. Changing diagram labels or structure needs a new generation unless a
separate supported editing capability exists; do not claim that this script edits
SmartArt node content. Recompose into a new output and render again.
