# -*- coding: utf-8 -*-
import os
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set


def doc_info(doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    return {
        "name": com_property(doc, "Name", ""),
        "full_name": com_property(doc, "FullName", ""),
        "paragraph_count": com_property(doc.Paragraphs, "Count", 0),
        "table_count": com_property(doc.Tables, "Count", 0),
        "section_count": com_property(doc.Sections, "Count", 1),
        "saved": com_property(doc, "Saved", False),
        "track_revisions": com_property(doc, "TrackRevisions", False),
        "comments_count": com_property(doc.Comments, "Count", 0),
    }


def doc_list() -> List[Dict]:
    from .app import list_documents
    return list_documents()


def doc_open(filepath: str) -> Dict:
    app = get_app()
    doc = app.Documents.Open(filepath)
    return {"name": doc.Name, "full_name": doc.FullName, "paragraph_count": doc.Paragraphs.Count, "table_count": doc.Tables.Count}


def doc_create(filename: Optional[str] = None) -> Dict:
    app = get_app()
    doc = app.Documents.Add()
    if filename:
        doc.SaveAs(filename)
    return {"name": doc.Name, "full_name": doc.FullName}


def doc_save(doc_index: Optional[int] = None, filepath: Optional[str] = None) -> Dict:
    doc = get_doc(doc_index)
    if filepath:
        doc.SaveAs(filepath)
    else:
        doc.Save()
    return {"name": com_property(doc, "Name", ""), "full_name": com_property(doc, "FullName", ""), "saved": True}


def doc_close(doc_index: Optional[int] = None, save_changes: bool = False) -> Dict:
    doc = get_doc(doc_index)
    name = com_property(doc, "Name", "")
    doc.Close(save_changes)
    return {"closed": name}


def doc_activate(doc_index: int) -> Dict:
    app = get_app()
    doc = app.Documents.Item(doc_index)
    doc.Activate()
    return {"active": doc.Name}


def doc_export_pdf(doc_index: Optional[int] = None, output_path: Optional[str] = None) -> Dict:
    doc = get_doc(doc_index)
    if output_path is None:
        base = os.path.splitext(doc.FullName or doc.Name)[0]
        output_path = base + ".pdf"
    wdFormatPDF = 17
    doc.SaveAs(output_path, FileFormat=wdFormatPDF)
    return {"pdf_path": output_path, "size": os.path.getsize(output_path) if os.path.exists(output_path) else 0}


# ====== Phase 5: Advanced Features ======

def insert_image(filepath: str, width: Optional[float] = None, height: Optional[float] = None,
                 position: str = "end", doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    if position == "end":
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    elif isinstance(position, int):
        rng = doc.Paragraphs.Item(position).Range
    else:
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    try:
        shape = doc.InlineShapes.AddPicture(filepath, 0, 1, com_property(rng, "Start", 0))
        if width:
            com_set(shape, "Width", width)
        if height:
            com_set(shape, "Height", height)
        return {"image": filepath, "width": width, "height": height, "inserted": True}
    except Exception as e:
        return {"error": str(e), "image": filepath}


def add_footnote(para_index: Optional[int] = None, text: str = "",
                 doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    try:
        if para_index:
            rng = doc.Paragraphs.Item(para_index).Range
        else:
            rng = get_app().Selection.Range
        note = doc.Footnotes.Add(rng, None, text)
        return {"footnote_index": note.Index, "text": text, "added": True}
    except Exception as e:
        # WPS COM footnote may fail on some document types; try alternative approach
        try:
            if para_index:
                rng = doc.Paragraphs.Item(para_index).Range
            else:
                rng = get_app().Selection.Range
            note = doc.Footnotes.Add(rng)
            if text:
                try:
                    note.Range.Text = text
                except Exception:
                    pass
            return {"footnote_index": note.Index, "text": text, "added": True}
        except Exception as e2:
            return {"error": str(e2), "footnote_added": False}


def add_endnote(para_index: Optional[int] = None, text: str = "",
                doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    try:
        if para_index:
            rng = doc.Paragraphs.Item(para_index).Range
        else:
            rng = get_app().Selection.Range
        note = doc.Endnotes.Add(rng, None, text)
        return {"endnote_index": note.Index, "text": text, "added": True}
    except Exception as e:
        try:
            if para_index:
                rng = doc.Paragraphs.Item(para_index).Range
            else:
                rng = get_app().Selection.Range
            note = doc.Endnotes.Add(rng)
            if text:
                try:
                    note.Range.Text = text
                except Exception:
                    pass
            return {"endnote_index": note.Index, "text": text, "added": True}
        except Exception as e2:
            return {"error": str(e2), "endnote_added": False}


def list_footnotes(doc_index: Optional[int] = None) -> List[Dict]:
    doc = get_doc(doc_index)
    result = []
    for i in range(1, doc.Footnotes.Count + 1):
        try:
            n = doc.Footnotes.Item(i)
            result.append({"index": i, "text": com_property(n.Range, "Text", "")[:100]})
        except Exception:
            continue
    return result


def add_bookmark(name: str, para_index: Optional[int] = None,
                 doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    if para_index:
        rng = doc.Paragraphs.Item(para_index).Range
    else:
        rng = get_app().Selection.Range
    doc.Bookmarks.Add(name, rng)
    return {"bookmark": name, "added": True}


def goto_bookmark(name: str, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    try:
        bm = doc.Bookmarks.Item(name)
        bm.Select()
        return {"bookmark": name, "found": True, "text": com_property(bm.Range, "Text", "")[:80]}
    except Exception:
        return {"bookmark": name, "found": False}


def list_bookmarks(doc_index: Optional[int] = None) -> List[Dict]:
    doc = get_doc(doc_index)
    result = []
    for i in range(1, doc.Bookmarks.Count + 1):
        try:
            bm = doc.Bookmarks.Item(i)
            result.append({"name": com_property(bm, "Name", ""), "text": com_property(bm.Range, "Text", "")[:60]})
        except Exception:
            continue
    return result


def insert_field(para_index: Optional[int] = None, field_code: str = "PAGE",
                 doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    if para_index:
        rng = doc.Paragraphs.Item(para_index).Range
    else:
        rng = get_app().Selection.Range
    field = doc.Fields.Add(rng, -1, field_code, True)
    return {"field_code": field_code, "result": com_property(field.Result, "Text", "")}


def add_watermark(text: str, font_size: float = 72, color: int = 15,
                  doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    try:
        for sec in range(1, doc.Sections.Count + 1):
            header = doc.Sections.Item(sec).Headers.Item(1)
            shp = header.Shapes.AddTextEffect(0, text, "宋体", font_size, 0, 0, 0, 0)
            shp.Fill.Visible = True
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = color if color > 0xFF else (color * 0x10000 + color * 0x100 + color)
            shp.Line.Visible = False
        return {"watermark": text, "added": True}
    except Exception as e:
        return {"error": str(e), "watermark": text}


def remove_watermark(doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    removed = 0
    # Remove from headers
    for sec in range(1, doc.Sections.Count + 1):
        try:
            header = doc.Sections.Item(sec).Headers.Item(1)
            for i in range(header.Shapes.Count, 0, -1):
                try:
                    header.Shapes.Item(i).Delete()
                    removed += 1
                except Exception:
                    continue
        except Exception:
            continue
    # Also remove TextEffect shapes from document body
    try:
        for i in range(doc.Shapes.Count, 0, -1):
            try:
                shp = doc.Shapes.Item(i)
                # Only delete shape if it looks like a watermark (TextEffect / WordArt)
                shp_type = com_property(shp, "Type", 0)
                if shp_type == 15:  # msoTextEffect
                    shp.Delete()
                    removed += 1
            except Exception:
                continue
    except Exception:
        pass
    return {"removed": removed}


def doc_properties(doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    props = {}
    for prop_name in ["Author", "Title", "Subject", "Keywords", "Category", "Comments",
                       "LastAuthor", "RevisionNumber", "TotalEditingTime", "CreationDate"]:
        try:
            val = com_property(doc.BuiltInDocumentProperties, prop_name, "")
            # Convert COM dates
            if hasattr(val, 'strftime'):
                val = val.strftime("%Y-%m-%d %H:%M:%S")
            props[prop_name] = val
        except Exception:
            continue
    props["page_count"] = com_property(doc.BuiltInDocumentProperties, "Number of pages", 0) if hasattr(doc.BuiltInDocumentProperties, "Number of pages") else 0
    return props


def set_doc_properties(author: Optional[str] = None, title: Optional[str] = None,
                       subject: Optional[str] = None, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    props = doc.BuiltInDocumentProperties
    updated = {}
    if author:
        com_set(props, "Author", author); updated["author"] = author
    if title:
        com_set(props, "Title", title); updated["title"] = title
    if subject:
        com_set(props, "Subject", subject); updated["subject"] = subject
    return {"updated": updated, "saved": com_property(doc, "Saved", False)}
