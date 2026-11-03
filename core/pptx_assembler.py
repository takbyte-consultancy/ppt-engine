"""Utilities for assembling saved exhibit slides into a new presentation."""
from __future__ import annotations

import copy
import io
import re

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn


_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _part_template(partname) -> str:
    path = str(partname)
    folder, filename = path.rsplit("/", 1)
    stem, dot, suffix = filename.rpartition(".")
    stem = re.sub(r"\d+$", "", stem) or "part"
    return f"{folder}/{stem}_import%d{dot}{suffix}" if dot else f"{folder}/{stem}_import%d"


def _copy_part(source_part, target_package, copied: dict, reserved: set[str]):
    """Copy an OOXML part graph (charts, workbooks, styles, and their relationships)."""
    if source_part in copied:
        return copied[source_part]
    template = _part_template(source_part.partname)
    new_partname = target_package.next_partname(template)
    suffix = 1
    while str(new_partname) in reserved:
        new_partname = PackURI(template % suffix)
        suffix += 1
    reserved.add(str(new_partname))
    target_part = PartFactory(new_partname, source_part.content_type, target_package, source_part.blob)
    copied[source_part] = target_part
    rid_map = {}
    for rel in source_part.rels.values():
        if rel.is_external:
            new_rid = target_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            related = _copy_part(rel.target_part, target_package, copied, reserved)
            new_rid = target_part.relate_to(related, rel.reltype)
        rid_map[rel.rId] = new_rid
    if isinstance(target_part, XmlPart):
        for element in target_part._element.iter():
            for attr_name, old_rid in list(element.attrib.items()):
                if attr_name.startswith(_REL_NS) and old_rid in rid_map:
                    element.set(attr_name, rid_map[old_rid])
    return target_part


def _fit_shape(shape, source_width: int, source_height: int,
               target_width: int, target_height: int) -> None:
    """Scale and center a shape while preserving the source slide's aspect ratio."""
    if any(value is None for value in (shape.left, shape.top, shape.width, shape.height)):
        return
    scale = min(target_width / source_width, target_height / source_height)
    offset_x = (target_width - source_width * scale) / 2
    offset_y = (target_height - source_height * scale) / 2
    shape.left = round(shape.left * scale + offset_x)
    shape.top = round(shape.top * scale + offset_y)
    shape.width = round(shape.width * scale)
    shape.height = round(shape.height * scale)
    # Explicit run sizes are in hundredths of a point and do not follow shape
    # geometry automatically when fitting a slide to the pitchbook canvas.
    for node in shape.element.iter():
        size = node.get("sz")
        if size and node.tag.rsplit("}", 1)[-1] in {"rPr", "defRPr", "endParaRPr", "latin"}:
            node.set("sz", str(max(100, round(int(size) * scale))))


def _attach_source_layout(source_slide, target_slide, output, copied_parts: dict,
                          reserved_parts: set[str]) -> None:
    """Import the source slide's layout/master/theme so inherited styling remains intact."""
    target_package = target_slide.part.package
    source_layout_part = source_slide.slide_layout.part
    target_layout_part = _copy_part(source_layout_part, target_package, copied_parts, reserved_parts)
    source_master_part = source_slide.slide_layout.slide_master.part
    target_master_part = _copy_part(source_master_part, target_package, copied_parts, reserved_parts)

    presentation_part = output.part
    master_rel = next((rel for rel in presentation_part.rels.values()
                       if not rel.is_external and rel.reltype == RT.SLIDE_MASTER
                       and rel.target_part is target_master_part), None)
    if master_rel is None:
        master_rid = presentation_part.relate_to(target_master_part, RT.SLIDE_MASTER)
        master_list = presentation_part._element.get_or_add_sldMasterIdLst()
        entry = OxmlElement("p:sldMasterId")
        used_ids = [int(item.get("id")) for item in master_list]
        entry.set("id", str(max([2147483648, *used_ids]) + 1))
        entry.set(qn("r:id"), master_rid)
        master_list.append(entry)

    for rel in list(target_slide.part.rels.values()):
        if rel.reltype == RT.SLIDE_LAYOUT:
            target_slide.part.drop_rel(rel.rId)
    target_slide.part.relate_to(target_layout_part, RT.SLIDE_LAYOUT)


def _copy_slide(source_slide, target_slide, output, copied_parts: dict,
                reserved_parts: set[str],
                source_width: int, source_height: int, target_width: int, target_height: int) -> None:
    """Copy editable shapes and their image references, preserving the source geometry/styles."""
    _attach_source_layout(source_slide, target_slide, output, copied_parts, reserved_parts)
    # Slide-level backgrounds and color-map overrides are independent of the
    # master/layout. Preserve them too, or theme colors can resolve differently.
    source_c_sld, target_c_sld = source_slide._element.cSld, target_slide._element.cSld
    if source_c_sld.bg is not None:
        if target_c_sld.bg is not None:
            target_c_sld.remove(target_c_sld.bg)
        target_c_sld.insert(0, copy.deepcopy(source_c_sld.bg))
    source_color_map = source_slide._element.clrMapOvr
    target_color_map = target_slide._element.clrMapOvr
    if source_color_map is not None:
        if target_color_map is not None:
            target_slide._element.remove(target_color_map)
        target_slide._element.append(copy.deepcopy(source_color_map))
    for shape in source_slide.shapes:
        clone = copy.deepcopy(shape.element)
        for element in clone.iter():
            for attr in ("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
                         "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
                         "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link"):
                old_rid = element.get(attr)
                if not old_rid:
                    continue
                rel = source_slide.part.rels[old_rid]
                if rel.is_external:
                    new_rid = target_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.IMAGE:
                    image_part = target_slide.part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))[0]
                    reserved_parts.add(str(image_part.partname))
                    new_rid = target_slide.part.relate_to(image_part, rel.reltype)
                else:
                    related_part = _copy_part(rel.target_part, target_slide.part.package, copied_parts, reserved_parts)
                    new_rid = target_slide.part.relate_to(related_part, rel.reltype)
                element.set(attr, new_rid)
        target_slide.shapes._spTree.insert_element_before(clone, "p:extLst")
        _fit_shape(target_slide.shapes[-1], source_width, source_height, target_width, target_height)


def assemble_exhibits(exhibits: list[tuple[dict, bytes]], target_size: tuple[int, int] | None = None) -> bytes:
    """Build a new PPTX from (provenance, one-slide PPTX bytes) entries in requested order."""
    if not exhibits:
        raise ValueError("A pitchbook must contain at least one exhibit slide")
    output = Presentation(io.BytesIO(exhibits[0][1]))
    if len(output.slides) != 1:
        raise ValueError("Each exhibit artifact must contain exactly one slide")
    source_width, source_height = output.slide_width, output.slide_height
    target_width, target_height = target_size or (source_width, source_height)
    if target_width <= 0 or target_height <= 0:
        raise ValueError("Pitchbook page dimensions must be positive")
    if (target_width, target_height) != (source_width, source_height):
        for shape in output.slides[0].shapes:
            _fit_shape(shape, source_width, source_height, target_width, target_height)
        output.slide_width, output.slide_height = target_width, target_height

    # Exhibit PPTX files can contain a reduced/custom layout collection. Do not
    # assume the default Office "Blank" layout is always at index 6.
    layouts = list(output.slide_layouts)
    if not layouts:
        raise ValueError("The first exhibit presentation has no slide layouts")
    blank = next((layout for layout in layouts if len(layout.placeholders) == 0), layouts[0])

    copied_parts = {}
    reserved_parts = {str(part.partname) for part in output.part.package.iter_parts()}
    for _, payload in exhibits[1:]:
        source = Presentation(io.BytesIO(payload))
        if len(source.slides) != 1:
            raise ValueError("Each exhibit artifact must contain exactly one slide")
        _copy_slide(source.slides[0], output.slides.add_slide(blank), output, copied_parts, reserved_parts,
                    source.slide_width, source.slide_height, target_width, target_height)

    for slide, (reference, _) in zip(output.slides, exhibits):
        slide._element.cSld.set("name", f"{reference.get('exhibit_name', reference.get('key', 'slide'))}__v{reference.get('version', '')}")

    stream = io.BytesIO()
    output.save(stream)
    return stream.getvalue()


def slide_components(pptx_bytes: bytes) -> list[list[dict]]:
    """Return stable-enough editable text component locators for feedback forms."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    rows = []
    for index, slide in enumerate(prs.slides, 1):
        rows.append([{"slide_number": index, "shape_id": shape.shape_id, "name": shape.name,
                      "text": shape.text[:300]}
                     for shape in slide.shapes
                     if getattr(shape, "has_text_frame", False) and shape.has_text_frame and shape.text.strip()])
    return rows
