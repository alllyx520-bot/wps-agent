# -*- coding: utf-8 -*-
"""Intelligence module — AI and rule-based document analysis and formatting."""
from .chinese_rules import CHINESE_FORMATTING
from .design_rules import (
    PPTX_PALETTES, XLSX_FINANCIAL_COLORS, PPTX_TYPOGRAPHY,
    PPTX_ANTI_PATTERNS, suggest_palette, suggest_typography,
    PPTX_SIZES, PPTX_SHAPE_TYPES, PPTX_CHART_TYPES,
    XLSX_NUMBER_FORMATS, XLSX_CELL_ROLE,
    DOCX_DXA, DOCX_TABLE_RULES, DOCX_PAGE_MARGINS,
    CN_FONTS, CN_FONT_SIZE_MAP, SMART_QUOTES,
)

__all__ = [
    "CHINESE_FORMATTING",
    "PPTX_PALETTES", "XLSX_FINANCIAL_COLORS", "PPTX_TYPOGRAPHY",
    "PPTX_ANTI_PATTERNS", "suggest_palette", "suggest_typography",
    "PPTX_SIZES", "PPTX_SHAPE_TYPES", "PPTX_CHART_TYPES",
    "XLSX_NUMBER_FORMATS", "XLSX_CELL_ROLE",
    "DOCX_DXA", "DOCX_TABLE_RULES", "DOCX_PAGE_MARGINS",
    "CN_FONTS", "CN_FONT_SIZE_MAP", "SMART_QUOTES",
]

try:
    from .llm_client import chat, analyze_document_structure, suggest_formatting, parse_natural_language_instructions
    __all__.extend(["chat", "analyze_document_structure", "suggest_formatting", "parse_natural_language_instructions"])
except ImportError:
    chat = analyze_document_structure = suggest_formatting = parse_natural_language_instructions = None

try:
    from .format_suggester import suggest
    __all__.append("suggest")
except ImportError:
    suggest = None

try:
    from .layout_analyzer import analyze as analyze_layout, generate_reformat_actions
    __all__.extend(["analyze_layout", "generate_reformat_actions"])
except ImportError:
    analyze_layout = generate_reformat_actions = None

try:
    from .format_intelligence import (
        auto_enhance, analyze_format_consistency, detect_document_type, suggest_format_fixes,
        auto_fix_format_issues, detect_paragraph_role, batch_detect_roles,
        format_health_report,
    )
    __all__.extend([
        "auto_enhance", "analyze_format_consistency", "detect_document_type", "suggest_format_fixes",
        "auto_fix_format_issues", "detect_paragraph_role", "batch_detect_roles",
        "format_health_report",
    ])
except ImportError as e:
    auto_enhance = analyze_format_consistency = detect_document_type = suggest_format_fixes = None
    auto_fix_format_issues = detect_paragraph_role = batch_detect_roles = None
    format_health_report = None

try:
    from .quality_supervisor import evaluate as evaluate_quality, sanitize_and_fix
    __all__.extend(["evaluate_quality", "sanitize_and_fix"])
except ImportError:
    evaluate_quality = sanitize_and_fix = None

try:
    from .template_manager import extract, save, load, list_all, delete, export_template, import_template, compare_with_template
    __all__.extend(["extract", "save", "load", "list_all", "delete", "export_template", "import_template", "compare_with_template"])
except ImportError:
    extract = save = load = list_all = delete = export_template = import_template = compare_with_template = None

try:
    from .content_generator import generate_content, summarize_document, rewrite_paragraph, expand_section, translate_section
    __all__.extend(["generate_content", "summarize_document", "rewrite_paragraph", "expand_section", "translate_section"])
except ImportError:
    generate_content = summarize_document = rewrite_paragraph = expand_section = translate_section = None