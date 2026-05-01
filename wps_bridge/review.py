# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set


def track_changes_toggle(enable, doc_index=None):
    doc = get_doc(doc_index)
    doc.TrackRevisions = enable
    return {"track_changes": bool(enable)}


def track_changes_status(doc_index=None):
    return {"track_changes": bool(com_property(get_doc(doc_index), "TrackRevisions", False))}


def comments_list(doc_index=None):
    comments = get_doc(doc_index).Comments
    result = []
    for i in range(1, comments.Count + 1):
        c = comments.Item(i)
        result.append({"index": i, "text": com_property(c.Range, "Text", "").strip(), "author": com_property(c, "Author", ""), "date": str(com_property(c, "Date", "")), "scope_text": com_property(c.Scope, "Text", "")[:80] if com_property(c, "Scope", None) else ""})
    return result


def comment_add(text, para_index=None, range_start=None, range_end=None, doc_index=None):
    doc = get_doc(doc_index)
    try:
        if para_index is not None:
            r = doc.Paragraphs.Item(para_index).Range
        elif range_start is not None and range_end is not None:
            r = doc.Range(range_start, range_end)
        else:
            r = get_app().Selection.Range
        doc.Comments.Add(r, text)
        return {"comment_added": text[:60], "added": True}
    except Exception as e:
        # WPS COM Comment.Add may fail if Track Changes is off or doc is protected
        # Try enabling track changes first then adding comment
        try:
            was_tracking = doc.TrackRevisions
            doc.Comments.Add(doc.Paragraphs.Item(para_index if para_index else 1).Range, text)
            if not was_tracking:
                doc.TrackRevisions = was_tracking
            return {"comment_added": text[:60], "added": True}
        except Exception as e2:
            return {"error": str(e2), "comment_added": False}


def revisions_list(doc_index=None):
    revs = get_doc(doc_index).Revisions
    result = []
    for i in range(1, revs.Count + 1):
        r = revs.Item(i)
        result.append({"index": i, "type": com_property(r, "Type", 0), "author": com_property(r, "Author", ""), "date": str(com_property(r, "Date", ""))})
    return result


def revisions_accept_all(doc_index=None):
    get_doc(doc_index).Revisions.AcceptAll()
    return {"accepted_all": True}


def revisions_reject_all(doc_index=None):
    get_doc(doc_index).Revisions.RejectAll()
    return {"rejected_all": True}
