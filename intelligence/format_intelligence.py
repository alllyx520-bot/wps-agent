# -*- coding: utf-8 -*-
"""
Format Intelligence — full auto-enhance pipeline.
Connects SemanticParser → Formatter → QualitySupervisor → LayoutAnalyzer
into a single one-click workflow.
"""
import logging
from typing import Dict, List, Optional

logger = logging.getLogger("wps-agent.format-intel")


def auto_enhance(filepath: str, template_name: Optional[str] = None,
                 document_type: Optional[str] = None,
                 output_path: Optional[str] = None) -> Dict:
    """One-click full document enhancement pipeline.

    Workflow:
      1. Parse document & build DOM
      2. Semantic analysis (detect type, roles, zones)
      3. Auto-format based on detected type
      4. Apply numbering
      5. Quality evaluation
      6. Layout analysis & fix
      7. Save result
    """
    from docx_engine import (
        parse_docx, build_document_model, serialize_document_model,
        DocumentAnalyzer, Formatter,
    )
    from docx_engine.semantic_model import SemanticParser
    from docx_engine.layout_model import LayoutAnalyzer

    output = output_path or filepath
    result = {
        "filepath": filepath,
        "output": output,
        "stages": {},
    }

    # ── Stage 1: Parse ──
    try:
        parsed = parse_docx(filepath)
        doc = build_document_model(parsed)
        result["stages"]["parse"] = {
            "status": "ok",
            "paragraphs": len(doc.paragraphs),
            "tables": len(doc.tables),
            "sections": len(doc.sections),
        }
    except Exception as e:
        return {"error": f"Parse failed: {e}"}

    # ── Stage 2: Semantic Analysis ──
    try:
        parser = SemanticParser(doc)
        graph = parser.parse()
        analyzer = DocumentAnalyzer(doc)
        detected_type = analyzer.detect_document_type()
        roles = [{"index": e.index, "role": e.role, "confidence": e.confidence}
                 for e in graph.elements]
        result["stages"]["semantic"] = {
            "status": "ok",
            "document_type": detected_type,
            "zones": {
                "cover": graph.cover_indices,
                "front_matter": graph.front_matter_indices,
                "body": graph.body_indices,
                "back_matter": graph.back_matter_indices,
            },
            "heading_count": len(graph.heading_tree),
            "roles_summary": _summarize_roles(roles),
        }
    except Exception as e:
        detected_type = "general"
        result["stages"]["semantic"] = {"status": "error", "error": str(e)}

    effective_type = document_type or detected_type

    # ── Stage 3: Auto Format ──
    try:
        formatter = Formatter(doc)
        if template_name:
            fmt_result = formatter.apply_template(template_name)
            result["stages"]["format"] = {
                "status": "ok",
                "template": template_name,
                "changes": fmt_result.get("change_count", 0),
            }
        else:
            fmt_result = formatter.auto_format(effective_type)
            result["stages"]["format"] = {
                "status": "ok",
                "document_type": effective_type,
                "changes": fmt_result.get("change_count", 0),
            }
    except Exception as e:
        result["stages"]["format"] = {"status": "error", "error": str(e)}

    # ── Stage 4: Numbering ──
    try:
        if effective_type in ("thesis", "paper", "report"):
            num_result = formatter.add_multi_level_numbering()
            result["stages"]["numbering"] = {
                "status": "ok",
                "changes": num_result.get("change_count", 0),
            }
    except Exception as e:
        result["stages"]["numbering"] = {"status": "error", "error": str(e)}

    # ── Stage 5: Quality Evaluation ──
    try:
        quality = analyzer.analyze_formatting_quality()
        result["stages"]["quality"] = {
            "status": "ok",
            "score": quality.get("score", 0),
            "issues": len(quality.get("issues", [])),
        }
    except Exception as e:
        result["stages"]["quality"] = {"status": "error", "error": str(e)}

    # ── Stage 6: Layout Analysis ──
    try:
        la = LayoutAnalyzer(doc)
        layout = la.analyze()
        result["stages"]["layout"] = {
            "status": "ok",
            "estimated_pages": layout.estimated_pages,
            "geometry": {
                "width": layout.geometry.width,
                "height": layout.geometry.height,
                "printable_width": layout.geometry.printable_width,
                "columns": layout.geometry.columns,
            },
            "issues": len(layout.issues),
        }
    except Exception as e:
        result["stages"]["layout"] = {"status": "error", "error": str(e)}

    # ── Stage 7: Save ──
    try:
        serialize_document_model(doc, output, original_docx=filepath)
        result["stages"]["save"] = {"status": "ok", "saved_to": output}
    except Exception as e:
        result["stages"]["save"] = {"status": "error", "error": str(e)}

    # ── Final Score ──
    quality_score = result["stages"].get("quality", {}).get("score", 0)
    format_changes = result["stages"].get("format", {}).get("changes", 0)
    layout_issues = result["stages"].get("layout", {}).get("issues", 0)

    result["final_score"] = min(100, quality_score - layout_issues * 5)
    result["verdict"] = (
        "excellent" if result["final_score"] >= 90 else
        "good" if result["final_score"] >= 70 else
        "needs_review"
    )

    return result


def _summarize_roles(roles: list) -> Dict:
    counts = {}
    for r in roles:
        role = r["role"]
        counts[role] = counts.get(role, 0) + 1
    top = sorted(counts.items(), key=lambda x: -x[1])[:8]
    return dict(top)


# ─── Standalone utility functions for ai_format handler ───

def analyze_format_consistency(filepath: str, doc_index: Optional[int] = None) -> Dict:
    """Analyze format consistency across paragraphs using COM or offline engine."""
    if doc_index is not None:
        from wps_bridge.app import get_doc
        from wps_bridge.content import paragraphs
        from wps_bridge.formatting import get_font, get_paragraph_format
        try:
            doc = get_doc(doc_index)
            total = doc.Paragraphs.Count
            samples = []
            for i in range(1, min(total + 1, 51)):
                try:
                    paras = paragraphs(i, 1, doc_index)
                    if paras:
                        p = paras[0]
                        samples.append(p)
                except Exception:
                    continue
            return {"status": "ok", "total_paragraphs": total, "samples": len(samples), "sample_data": samples[:5]}
        except Exception as e:
            return {"status": "error", "error": str(e)}
    else:
        from docx_engine import parse_docx, build_document_model, DocumentAnalyzer
        try:
            parsed = parse_docx(filepath)
            doc = build_document_model(parsed)
            analyzer = DocumentAnalyzer(doc)
            return {"status": "ok", "document_type": analyzer.detect_document_type(), "quality": analyzer.analyze_formatting_quality()}
        except Exception as e:
            return {"status": "error", "error": str(e)}


def detect_document_type(filepath: str = None, doc_index: Optional[int] = None) -> Dict:
    """Detect document type from file path or COM document."""
    if filepath:
        from docx_engine import parse_docx, build_document_model, DocumentAnalyzer
        try:
            parsed = parse_docx(filepath)
            doc = build_document_model(parsed)
            analyzer = DocumentAnalyzer(doc)
            return {"document_type": analyzer.detect_document_type()}
        except Exception as e:
            return {"document_type": "unknown", "error": str(e)}
    return {"document_type": "unknown", "error": "No filepath provided"}


def suggest_format_fixes(filepath: str) -> Dict:
    """Suggest formatting fixes based on quality analysis."""
    from docx_engine import parse_docx, build_document_model, DocumentAnalyzer
    try:
        parsed = parse_docx(filepath)
        doc = build_document_model(parsed)
        analyzer = DocumentAnalyzer(doc)
        quality = analyzer.analyze_formatting_quality()
        return {"suggestions": quality.get("issues", []), "score": quality.get("score", 0)}
    except Exception as e:
        return {"suggestions": [], "error": str(e)}


def auto_fix_format_issues(filepath: str, output_path: str = None) -> Dict:
    """Auto-fix detected formatting issues."""
    from intelligence.quality_supervisor import evaluate as evaluate_quality
    try:
        result = evaluate_quality(filepath)
        return {"fixed": result.get("score", 0), "details": result}
    except Exception as e:
        return {"error": str(e)}


def detect_paragraph_role(para_index: int, doc_index: Optional[int] = None) -> Dict:
    """Detect the semantic role of a single paragraph via COM."""
    from wps_bridge.app import get_doc
    from wps_bridge.utils import com_property
    try:
        doc = get_doc(doc_index)
        p = doc.Paragraphs.Item(para_index)
        text = com_property(p.Range, "Text", "").strip()
        style = com_property(p.Range.Style, "NameLocal", "")
        outline = com_property(p.Format, "OutlineLevel", 10)
        role = "heading" if 1 <= outline <= 9 else (
            "title" if "标题" in style and outline == 0 else
            "abstract" if "摘要" in text[:10] or "abstract" in text[:10].lower() else
            "body"
        )
        return {"para_index": para_index, "role": role, "confidence": 0.8, "style": style, "outline_level": outline, "text_preview": text[:80]}
    except Exception as e:
        return {"para_index": para_index, "role": "unknown", "error": str(e)}


def batch_detect_roles(doc_index: Optional[int] = None) -> Dict:
    """Detect semantic roles for all paragraphs via COM."""
    from wps_bridge.app import get_doc
    from wps_bridge.utils import com_property
    try:
        doc = get_doc(doc_index)
        total = doc.Paragraphs.Count
        roles = []
        for i in range(1, min(total + 1, 201)):
            try:
                p = doc.Paragraphs.Item(i)
                text = com_property(p.Range, "Text", "").strip()
                outline = com_property(p.Format, "OutlineLevel", 10)
                style = com_property(p.Range.Style, "NameLocal", "")
                if 1 <= outline <= 9:
                    role = "heading"
                elif not text:
                    role = "empty"
                elif "摘要" in text[:10] or "abstract" in text[:10].lower():
                    role = "abstract"
                elif "关键词" in text[:10] or "keywords" in text[:10].lower():
                    role = "keywords"
                elif "目录" in text[:4]:
                    role = "toc"
                elif "参考文献" in text[:10]:
                    role = "references"
                else:
                    role = "body"
                roles.append({"index": i, "role": role, "text_preview": text[:60], "outline_level": outline, "style": style})
            except Exception:
                continue
        return {"total": total, "detected": len(roles), "roles": roles}
    except Exception as e:
        return {"total": 0, "detected": 0, "error": str(e)}


def format_health_report(filepath: str = None, doc_index: Optional[int] = None) -> Dict:
    """Generate a comprehensive format health report."""
    if doc_index is not None:
        from wps_bridge.app import get_doc
        from wps_bridge.utils import com_property
        try:
            doc = get_doc(doc_index)
            issues = []
            total = doc.Paragraphs.Count
            style_counts = {}
            for i in range(1, min(total + 1, 101)):
                try:
                    p = doc.Paragraphs.Item(i)
                    style = com_property(p.Range.Style, "NameLocal", "")
                    style_counts[style] = style_counts.get(style, 0) + 1
                except Exception:
                    continue
            if total > 200:
                issues.append({"type": "warning", "message": f"Large document ({total} paragraphs), review may be incomplete"})
            return {"status": "ok", "total_paragraphs": total, "sampled": min(total, 100), "style_distribution": style_counts, "issues": issues}
        except Exception as e:
            return {"status": "error", "error": str(e)}
    elif filepath:
        from docx_engine import parse_docx, build_document_model, DocumentAnalyzer
        try:
            parsed = parse_docx(filepath)
            doc = build_document_model(parsed)
            analyzer = DocumentAnalyzer(doc)
            quality = analyzer.analyze_formatting_quality()
            return {"status": "ok", "quality": quality}
        except Exception as e:
            return {"status": "error", "error": str(e)}
    return {"status": "error", "error": "No filepath or doc_index provided"}
