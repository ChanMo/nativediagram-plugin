# NativeDiagram for Claude

NativeDiagram creates native, editable Microsoft Office SmartArt from a process, hierarchy, cycle, or other supported structure. Its remote MCP connection finds layouts and produces a standalone Word document, a PowerPoint slide, and a PNG preview. The bundled skill helps Claude choose a suitable layout and, when local document tools are available, insert the generated SmartArt into a complete report or presentation.

## Connect

1. Install the NativeDiagram plugin in a supported Claude client.
2. Connect the `nativediagram` MCP server through Claude's OAuth flow using a [NativeDiagram account](https://nativediagram.com/).
3. Ask Claude to create an editable SmartArt diagram for Word or PowerPoint. For a complete document, provide the surrounding content and requested insertion point.

You need a verified NativeDiagram account and available credits. Listing templates and viewing up to three candidate sample images are free; each successful diagram creation uses one credit. Keep the request ID for safe recovery after a disconnected response: completed keyed requests return the original files without another charge. Generated file download links are valid for seven days and can be refreshed for free through result recovery. The files remain in your NativeDiagram account under the [privacy policy](https://nativediagram.com/privacy/).

The optional insertion script requires Python 3 and `lxml`. It runs locally, reads a generated file and a target file of the same Office format, and writes a new file. See the [composition guidance](skills/native-smartart/references/composition.md) for supported placeholders and limits. The surrounding document is created with the document tools available in your Claude environment.

The MCP receives the labels, relationships, and layout choice needed to make the diagram. It does not require the complete target document or conversation. For setup instructions, see the [connection guide](https://nativediagram.com/docs/connect/). For help, contact [hello@oddloaf.dev](mailto:hello@oddloaf.dev).

The plugin source is available under the [MIT License](LICENSE).

## Upgrade to 0.3.0

Refresh the MCP connection and replace the installed SmartArt skill to load the updated tools. Creation accepts a retained request UUID for safe recovery; calls without it remain new charged creations. Relationship diagrams now require an explicit overlap or equation choice. Radial diagrams default to neutral association; inward or outward arrows require an explicit choice. Retired templates can no longer be selected by ID. These stricter choices protect the meaning of your diagram.
