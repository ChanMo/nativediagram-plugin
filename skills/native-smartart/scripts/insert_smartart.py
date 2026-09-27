"""Insert one NativeDiagram artifact at an explicit DOCX/PPTX placeholder.

Local-only; Python 3 + lxml. Copies the diagram's dependency graph into fresh
parts, preserves unrelated target parts and inherits the target document theme.
Supports a single-section DOCX body paragraph or a top-level PPTX text shape.
Does not render, infer insertion positions, flatten SmartArt, or overwrite inputs.
"""
from __future__ import annotations

import argparse
import copy
import io
import posixpath
from pathlib import Path
import uuid
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as E

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'dsp': 'http://schemas.microsoft.com/office/drawing/2008/diagram',
}
CT = '[Content_Types].xml'
DIAGRAM_TYPES = {'diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors', 'diagramDrawing'}
SCALE_ATTRS = {'off': ('x', 'y'), 'chOff': ('x', 'y'), 'ext': ('cx', 'cy'),
               'chExt': ('cx', 'cy'), 'bodyPr': ('lIns', 'tIns', 'rIns', 'bIns'),
               'defRPr': ('sz',), 'rPr': ('sz',), 'endParaRPr': ('sz',), 'ln': ('w',)}
SCALE_ATTRS.update({f'lvl{i}pPr': ('marL', 'indent') for i in range(1, 10)})


def xml(payload):
    parser = E.XMLParser(resolve_entities=False, no_network=True)
    root = E.fromstring(payload, parser)
    if root.getroottree().docinfo.doctype:
        raise ValueError('DTD declarations are not supported.')
    return root


def serialize(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def read_package(data):
    with ZipFile(io.BytesIO(data)) as archive:
        entries = archive.infolist()
        if len(entries) > 4096 or sum(item.file_size for item in entries) > 100 * 1024 * 1024:
            raise ValueError('Package exceeds the supported size limit.')
        names = [item.filename for item in entries]
        if len(names) != len(set(names)) or any(name.startswith('_xmlsignatures/') for name in names):
            raise ValueError('Duplicate parts and signed packages are not supported.')
        return {name: archive.read(name) for name in names}


def rels_name(part):
    folder, name = posixpath.split(part)
    return posixpath.join(folder, '_rels', name + '.rels')


def relations(package, part):
    name = rels_name(part)
    return xml(package[name]) if name in package else E.Element('{%s}Relationships' % NS['rel'], nsmap={None: NS['rel']})


def resolve(part, target):
    name = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)) if not target.startswith('/') else target[1:]
    if name.startswith('../') or '\\' in name or '?' in name or '#' in name:
        raise ValueError('Unsupported relationship target.')
    return name


def single(items, description):
    if len(items) != 1:
        raise ValueError(f'Expected exactly one {description}; found {len(items)}.')
    return items[0]


def scale_cache(root, factor):
    for transform in root.xpath('./dsp:spTree/dsp:grpSpPr/a:xfrm', namespaces=NS):
        transform.getparent().remove(transform)
    for element in root.iter():
        for attr in SCALE_ATTRS.get(E.QName(element).localname, ()):
            if element.get(attr) is not None:
                element.set(attr, str(round(int(element.get(attr)) * factor)))


def insert_smartart(target_bytes, diagram_bytes, *, kind, marker, width_cm=None):
    """Return a new package, or fail before any output is written.

    DOCX marker is the exact text of an otherwise empty body paragraph. PPTX
    marker is the exact text of a top-level text box whose bounds define the slot.
    Source and target must have the same format. The target theme remains intact.
    """
    if not marker or kind not in ('docx', 'pptx'):
        raise ValueError('Specify a non-empty marker and docx or pptx format.')
    source, target = read_package(diagram_bytes), read_package(target_bytes)
    content_types = xml(target[CT])
    source_ct = xml(source[CT])
    overrides = {x.get('PartName').lstrip('/'): x.get('ContentType') for x in source_ct if E.QName(x).localname == 'Override'}
    defaults = {x.get('Extension'): x.get('ContentType') for x in source_ct if E.QName(x).localname == 'Default'}

    if kind == 'docx':
        owner = source_owner = 'word/document.xml'
        root, source_root = xml(target[owner]), xml(source[source_owner])
        sections = root.xpath('.//w:sectPr', namespaces=NS)
        section = single(sections, 'document section')
        if section.xpath('./w:cols[@w:num != "1"]', namespaces=NS):
            raise ValueError('Multi-column documents are not supported.')
        candidates = root.xpath('./w:body/w:p', namespaces=NS)
        placeholder = single([p for p in candidates if ''.join(p.xpath('.//w:t/text()', namespaces=NS)) == marker], 'body marker paragraph')
        if (any(child.tag not in ('{%s}pPr' % NS['w'], '{%s}r' % NS['w']) for child in placeholder)
                or placeholder.xpath('.//w:drawing | .//w:object | .//w:fldChar | .//w:br | .//w:sectPr', namespaces=NS)):
            raise ValueError('Marker paragraph must contain only marker text.')
        block = copy.deepcopy(single(source_root.xpath('.//w:drawing[.//dgm:relIds]', namespaces=NS), 'source SmartArt drawing'))
        inline = single(block.xpath('./wp:inline', namespaces=NS), 'inline source diagram')
        extent = single(inline.xpath('./wp:extent', namespaces=NS), 'source extent')
        sw, sh = int(extent.get('cx')), int(extent.get('cy'))
        page = single(section.xpath('./w:pgSz', namespaces=NS), 'page size')
        margin = single(section.xpath('./w:pgMar', namespaces=NS), 'page margins')
        usable = (int(page.get('{%s}w' % NS['w'])) - sum(int(margin.get('{%s}%s' % (NS['w'], key), '0')) for key in ('left', 'right', 'gutter'))) * 635
        width = round(width_cm * 360000) if width_cm is not None else min(sw, usable)
        if width <= 0 or width > usable or sw <= 0 or sh <= 0:
            raise ValueError('Diagram width must fit the positive page content width.')
        factor = width / sw
        extent.set('cx', str(width)); extent.set('cy', str(round(sh * factor)))
        existing_ids = []
        for name, payload in target.items():
            if name.startswith('word/') and name.endswith('.xml'):
                existing_ids.extend(int(v) for v in xml(payload).xpath('.//wp:docPr/@id', namespaces=NS))
        max_id = max(existing_ids or [0])
        for index, prop in enumerate(block.xpath('.//wp:docPr', namespaces=NS), max_id + 1):
            prop.set('id', str(index))
        # Reuse paragraph formatting while replacing only its marker runs.
        replacement = E.Element('{%s}p' % NS['w'])
        props = placeholder.find('{%s}pPr' % NS['w'])
        if props is not None:
            replacement.append(copy.deepcopy(props))
        run = E.SubElement(replacement, '{%s}r' % NS['w']); run.append(block)
    else:
        if width_cm is not None:
            raise ValueError('PPTX uses the marker text box bounds, not --width-cm.')
        slides = [name for name in target if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
        hits = []
        for name in slides:
            slide = xml(target[name])
            for shape in slide.xpath('./p:cSld/p:spTree/p:sp', namespaces=NS):
                if ''.join(shape.xpath('.//a:t/text()', namespaces=NS)) == marker:
                    hits.append((name, slide, shape))
        owner, root, placeholder = single(hits, 'slide marker text box')
        sources = []
        for name in source:
            if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
                sources.extend((name, frame) for frame in xml(source[name]).xpath('.//p:graphicFrame[.//dgm:relIds]', namespaces=NS))
        source_owner, frame = single(sources, 'source SmartArt frame')
        block = copy.deepcopy(frame)
        box = single(placeholder.xpath('./p:spPr/a:xfrm', namespaces=NS), 'explicit marker geometry')
        if box.get('rot', '0') != '0' or box.get('flipH', '0') not in ('0', 'false') or box.get('flipV', '0') not in ('0', 'false'):
            raise ValueError('Rotated or flipped placeholders are not supported.')
        old = single(block.xpath('./p:xfrm', namespaces=NS), 'source frame transform')
        sw, sh = int(old.find('a:ext', NS).get('cx')), int(old.find('a:ext', NS).get('cy'))
        bw, bh = int(box.find('a:ext', NS).get('cx')), int(box.find('a:ext', NS).get('cy'))
        if min(sw, sh, bw, bh) <= 0:
            raise ValueError('Frame dimensions must be positive.')
        factor = min(bw / sw, bh / sh)
        width, height = round(sw * factor), round(sh * factor)
        old.find('a:ext', NS).attrib.update({'cx': str(width), 'cy': str(height)})
        old.find('a:off', NS).attrib.update({'x': str(int(box.find('a:off', NS).get('x')) + (bw-width)//2), 'y': str(int(box.find('a:off', NS).get('y')) + (bh-height)//2)})
        prop = single(block.xpath('./p:nvGraphicFramePr/p:cNvPr', namespaces=NS), 'frame identity')
        prop.set('id', str(max([int(v) for v in root.xpath('.//p:cNvPr/@id', namespaces=NS)] or [0]) + 1))
        prop.set('name', 'NativeDiagram SmartArt')
        replacement = block

    source_rels = relations(source, source_owner)
    target_rels = relations(target, owner)
    top_map = {}
    selected = [r for r in source_rels if r.get('Type', '').rsplit('/', 1)[-1] in DIAGRAM_TYPES]
    for rel in selected:
        top_map[rel.get('Id')] = 'rIdND' + uuid.uuid4().hex
    for element in block.iter():
        for attr, value in list(element.attrib.items()):
            if attr.startswith('{%s}' % NS['r']):
                if value not in top_map:
                    raise ValueError('Source drawing references a non-diagram relationship.')
                element.set(attr, top_map[value])
    copied = {}
    prefix = 'word' if kind == 'docx' else 'ppt'

    def copy_part(name):
        if name in copied:
            return copied[name]
        if name not in source:
            raise ValueError(f'Missing source diagram dependency: {name}')
        new = f'{prefix}/diagrams/nd_{uuid.uuid4().hex}_{posixpath.basename(name)}'
        copied[name] = new
        mime = overrides.get(name) or defaults.get(name.rsplit('.', 1)[-1])
        if not mime:
            raise ValueError(f'Missing content type: {name}')
        E.SubElement(content_types, '{%s}Override' % NS['ct'], PartName='/' + new, ContentType=mime)
        payload = source[name]
        if name.endswith('.xml'):
            tree = xml(payload)
            # NativeDiagram data extensions can point at a drawing relation on
            # the document/slide rather than on the data part itself.
            local_ids = {r.get('Id') for r in relations(source, name)}
            for element in tree.iter():
                for attr, value in list(element.attrib.items()):
                    if (attr.startswith('{%s}' % NS['r']) or E.QName(attr).localname == 'relId') and value in top_map and value not in local_ids:
                        element.set(attr, top_map[value])
            if E.QName(tree).namespace == NS['dsp']:
                scale_cache(tree, factor)
            payload = serialize(tree)
        target[new] = payload
        nested = relations(source, name)
        for rel in nested:
            if rel.get('TargetMode') == 'External':
                raise ValueError('External diagram dependencies are not supported.')
            dependency = resolve(name, rel.get('Target', ''))
            rel.set('Target', posixpath.relpath(copy_part(dependency), posixpath.dirname(new)))
        if len(nested):
            target[rels_name(new)] = serialize(nested)
        return new

    for rel in selected:
        if rel.get('TargetMode') == 'External':
            raise ValueError('External diagram dependencies are not supported.')
        new = copy_part(resolve(source_owner, rel.get('Target', '')))
        E.SubElement(target_rels, '{%s}Relationship' % NS['rel'], Id=top_map[rel.get('Id')], Type=rel.get('Type'), Target=posixpath.relpath(new, posixpath.dirname(owner)))
    if not selected:
        raise ValueError('No native diagram relationships found.')
    placeholder.getparent().replace(placeholder, replacement)
    target[owner], target[rels_name(owner)], target[CT] = serialize(root), serialize(target_rels), serialize(content_types)
    output = io.BytesIO()
    with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
        for name, payload in target.items():
            archive.writestr(name, payload)
    return output.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--target', type=Path, required=True)
    parser.add_argument('--diagram', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--marker', required=True)
    parser.add_argument('--width-cm', type=float)
    args = parser.parse_args()
    if args.target.suffix.lower() != args.diagram.suffix.lower() or args.output.suffix.lower() != args.target.suffix.lower():
        parser.error('Target, diagram and output must use the same .docx or .pptx format.')
    result = insert_smartart(args.target.read_bytes(), args.diagram.read_bytes(), kind=args.target.suffix.lower()[1:], marker=args.marker, width_cm=args.width_cm)
    with args.output.open('xb') as file:
        file.write(result)
    print(f'Created {args.output}. Target theme retained. Render and inspect before delivery.')


if __name__ == '__main__':
    main()
