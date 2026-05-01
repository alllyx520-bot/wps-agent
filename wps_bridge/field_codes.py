# -*- coding: utf-8 -*-
"""Field Codes bridge — advanced field manipulation (Quote, If, Seq, StyleRef, Ref, DocProperty, etc.)."""
from typing import Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set

FIELD_CODES = {
    "PAGE": "wdFieldPage",
    "NUMPAGES": "wdFieldNumPages",
    "DATE": "wdFieldDate",
    "TIME": "wdFieldTime",
    "FILENAME": "wdFieldFileName",
    "AUTHOR": "wdFieldAuthor",
    "TITLE": "wdFieldTitle",
    "SUBJECT": "wdFieldSubject",
    "KEYWORDS": "wdFieldKeywords",
    "COMMENTS": "wdFieldComments",
    "LASTSAVEDBY": "wdFieldLastSavedBy",
    "DOCPROPERTY": "wdFieldDocProperty",
    "CREATEDATE": "wdFieldCreateDate",
    "SAVEDATE": "wdFieldSaveDate",
    "PRINTDATE": "wdFieldPrintDate",
    "FILESIZE": "wdFieldFileSize",
    "NUMWORDS": "wdFieldNumWords",
    "NUMCHARS": "wdFieldNumChars",
    "QUOTE": "wdFieldQuote",
    "IF": "wdFieldIf",
    "SEQ": "wdFieldSequence",
    "STYLEREF": "wdFieldStyleRef",
    "REF": "wdFieldRef",
    "PAGEREF": "wdFieldPageRef",
    "NOTEREF": "wdFieldNoteRef",
    "INCLUDEPICTURE": "wdFieldIncludePicture",
    "INCLUDETEXT": "wdFieldIncludeText",
    "HYPERLINK": "wdFieldHyperlink",
    "SECTION": "wdFieldSection",
    "SECTIONPAGES": "wdFieldSectionPages",
    "TOC": "wdFieldTOC",
    "TOA": "wdFieldTOA",
    "INDEX": "wdFieldIndex",
    "XE": "wdFieldIndexEntry",
    "TA": "wdFieldTOAEntry",
    "TC": "wdFieldTOCEntry",
    "EQ": "wdFieldEQ",
    "FORMULA": "wdFieldFormula",
    "SYMBOL": "wdFieldSymbol",
    "MERGEFIELD": "wdFieldMergeField",
    "NEXT": "wdFieldNext",
    "NEXTIF": "wdFieldNextIf",
    "SKIPIF": "wdFieldSkipIf",
    "FILLIN": "wdFieldFillIn",
    "ASK": "wdFieldAsk",
    "COMPARE": "wdFieldCompare",
    "DATABASE": "wdFieldDatabase",
}


def insert_field(field_code: str, switches: Optional[str] = None,
                 para_index: Optional[int] = None, position: str = "end",
                 doc_index: Optional[int] = None) -> Dict:
    """Insert a field code at the specified position.
    
    field_code: e.g. 'PAGE', 'DATE  \\@ \"yyyy-MM-dd\"', 'SEQ Figure \\* ARABIC'
    switches: additional field switches
    """
    doc = get_doc(doc_index)
    code = field_code
    if switches:
        code = f"{field_code} {switches}"

    if position == "end":
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    elif position == "before" and para_index:
        rng = doc.Paragraphs.Item(para_index).Range
    elif position == "after" and para_index:
        p_range = doc.Paragraphs.Item(para_index).Range
        rng = doc.Range(p_range.End - 1, p_range.End - 1)
    else:
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

    try:
        field = doc.Fields.Add(rng, -1, code, True)
        return {
            "inserted": True,
            "field_code": code,
            "position": position,
            "result_text": com_property(field.Result, "Text", ""),
        }
    except Exception as e:
        return {"error": f"Failed to insert field: {e}"}


def insert_quote(text: str, para_index: Optional[int] = None,
                 doc_index: Optional[int] = None) -> Dict:
    """Insert a QUOTE field with literal text."""
    return insert_field(f'QUOTE "{text}"', para_index=para_index, doc_index=doc_index)


def insert_doc_property(property_name: str, para_index: Optional[int] = None,
                        doc_index: Optional[int] = None) -> Dict:
    """Insert a DOCPROPERTY field referencing a document property."""
    return insert_field(f"DOCPROPERTY {property_name}", para_index=para_index, doc_index=doc_index)


def insert_seq(sequence_name: str, format_type: str = "ARABIC",
               para_index: Optional[int] = None,
               doc_index: Optional[int] = None) -> Dict:
    """Insert a SEQ (sequence) field for auto-numbering.
    sequence_name: e.g. 'Figure', 'Table', 'Equation'
    format_type: ARABIC, ROMAN, ALPHABETIC, etc.
    """
    return insert_field(f"SEQ {sequence_name} \\* {format_type}", para_index=para_index, doc_index=doc_index)


def insert_style_ref(style_name: str, switches: Optional[str] = None,
                     para_index: Optional[int] = None,
                     doc_index: Optional[int] = None) -> Dict:
    """Insert a STYLEREF field (e.g. for running headers showing chapter title).
    style_name: 'Heading 1', '标题 1', etc.
    """
    code = f"STYLEREF \"{style_name}\""
    if switches:
        code += f" {switches}"
    return insert_field(code, para_index=para_index, doc_index=doc_index)


def insert_ref(bookmark_name: str, switches: Optional[str] = None,
               para_index: Optional[int] = None,
               doc_index: Optional[int] = None) -> Dict:
    """Insert a REF field referencing a bookmark."""
    code = f"REF {bookmark_name}"
    if switches:
        code += f" {switches}"
    return insert_field(code, para_index=para_index, doc_index=doc_index)


def insert_if(condition: str, true_text: str, false_text: str,
              para_index: Optional[int] = None,
              doc_index: Optional[int] = None) -> Dict:
    """Insert an IF field with condition.
    condition: e.g. '{ MERGEFIELD Amount } > 1000'
    """
    return insert_field(
        f'IF {condition} "{true_text}" "{false_text}"',
        para_index=para_index, doc_index=doc_index,
    )


def list_fields(doc_index: Optional[int] = None) -> Dict:
    """List all fields in the document."""
    doc = get_doc(doc_index)
    result = []
    total = doc.Fields.Count
    for i in range(1, min(total + 1, 200)):
        try:
            f = doc.Fields.Item(i)
            result.append({
                "index": i,
                "code": com_property(f.Code, "Text", "").strip()[:120],
                "result": com_property(f.Result, "Text", "")[:60],
                "kind": com_property(f, "Kind", -1),
                "locked": com_property(f, "Locked", False),
            })
        except Exception:
            continue
    return {"total": total, "fields": result}


def update_fields(doc_index: Optional[int] = None) -> Dict:
    """Update all fields in the document."""
    doc = get_doc(doc_index)
    count = doc.Fields.Count
    for i in range(1, count + 1):
        try:
            doc.Fields.Item(i).Update()
        except Exception:
            continue
    return {"updated": count}


def unlink_field(field_index: int, doc_index: Optional[int] = None) -> Dict:
    """Unlink a field (convert to plain text)."""
    doc = get_doc(doc_index)
    f = doc.Fields.Item(field_index)
    text = com_property(f.Result, "Text", "")
    f.Unlink()
    return {"unlinked": True, "result_text": text[:80]}


def find_field_by_code(pattern: str, doc_index: Optional[int] = None) -> Dict:
    """Find fields whose code matches a pattern (substring search)."""
    doc = get_doc(doc_index)
    results = []
    total = doc.Fields.Count
    for i in range(1, total + 1):
        try:
            f = doc.Fields.Item(i)
            code = com_property(f.Code, "Text", "").strip()
            if pattern.lower() in code.lower():
                results.append({
                    "index": i,
                    "code": code[:120],
                    "result": com_property(f.Result, "Text", "")[:60],
                })
        except Exception:
            continue
    return {"pattern": pattern, "found": len(results), "fields": results}
