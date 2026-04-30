# -*- coding: utf-8 -*-
"""
Design rules and conventions from Claude official skills (docx/xlsx/pptx).
Injected into wps-agent intelligence layer for professional document creation.
All values are knowledge references, not hard rules — agent decides when to apply.
"""

# ============================================================
# PPTX Design Palettes (from pptx skill)
# ============================================================

PPTX_PALETTES = {
    "midnight_executive": {
        "primary": "1E2761", "secondary": "CADCFC", "accent": "FFFFFF",
        "name": "午夜行政",
        "style": "dark_header",
    },
    "forest_moss": {
        "primary": "2C5F2D", "secondary": "97BC62", "accent": "F5F5F5",
        "name": "森林苔藓",
        "style": "nature",
    },
    "coral_energy": {
        "primary": "F96167", "secondary": "F9E795", "accent": "2F3C7E",
        "name": "珊瑚能量",
        "style": "vibrant",
    },
    "warm_terracotta": {
        "primary": "B85042", "secondary": "E7E8D1", "accent": "A7BEAE",
        "name": "暖陶土",
        "style": "warm",
    },
    "ocean_gradient": {
        "primary": "065A82", "secondary": "1C7293", "accent": "21295C",
        "name": "海洋渐变",
        "style": "professional",
    },
    "charcoal_minimal": {
        "primary": "36454F", "secondary": "F2F2F2", "accent": "212121",
        "name": "炭灰极简",
        "style": "minimal",
    },
    "teal_trust": {
        "primary": "028090", "secondary": "00A896", "accent": "02C39A",
        "name": "青色信赖",
        "style": "fresh",
    },
    "berry_cream": {
        "primary": "6D2E46", "secondary": "A26769", "accent": "ECE2D0",
        "name": "莓果奶油",
        "style": "elegant",
    },
    "sage_calm": {
        "primary": "84B59F", "secondary": "69A297", "accent": "50808E",
        "name": "鼠尾草平静",
        "style": "calm",
    },
    "cherry_bold": {
        "primary": "990011", "secondary": "FCF6F5", "accent": "2F3C7E",
        "name": "樱桃大胆",
        "style": "bold",
    },
}

# ============================================================
# PPTX Typography Pairings (from pptx skill)
# ============================================================

PPTX_TYPOGRAPHY = [
    {"header": "Georgia", "body": "Calibri", "style": "classic"},
    {"header": "Arial Black", "body": "Arial", "style": "bold"},
    {"header": "Calibri", "body": "Calibri Light", "style": "modern"},
    {"header": "Cambria", "body": "Calibri", "style": "formal"},
    {"header": "Trebuchet MS", "body": "Calibri", "style": "friendly"},
    {"header": "Impact", "body": "Arial", "style": "strong"},
    {"header": "Palatino", "body": "Garamond", "style": "literary"},
    {"header": "Consolas", "body": "Calibri", "style": "technical"},
]

# Size hierarchy
PPTX_SIZES = {
    "title": {"min": 36, "max": 44, "bold": True, "unit": "pt"},
    "section": {"min": 20, "max": 24, "bold": True, "unit": "pt"},
    "body": {"min": 14, "max": 16, "bold": False, "unit": "pt"},
    "caption": {"min": 10, "max": 12, "bold": False, "unit": "pt", "muted": True},
}

PPTX_SPACING = {
    "margin_min": 0.5,  # inches
    "block_gap": "0.3-0.5 inches",
    "title_padding": 0,  # set margin=0 for title alignment
}

# ============================================================
# PPTX Design Anti-Patterns (10 rules from pptx skill)
# ============================================================

PPTX_ANTI_PATTERNS = [
    "Don't repeat the same layout — vary columns, cards, callouts across slides",
    "Don't center body text — left-align paragraphs; center only titles",
    "Don't skimp on size contrast — titles need 36pt+ to stand out from 14-16pt body",
    "Don't default to blue — pick colors reflecting the specific topic",
    "Don't mix spacing randomly — choose 0.3in or 0.5in gaps consistently",
    "Don't style one slide and leave rest plain — commit fully or keep simple",
    "Don't create text-only slides — add images, icons, charts, or shapes",
    "Don't forget text box padding — set margin:0 when aligning text with shapes",
    "Don't use low-contrast elements — icons AND text need strong contrast vs background",
    "Never use accent lines under titles — hallmark of AI-generated slides",
]

# ============================================================
# PPTX Shape Types
# ============================================================

PPTX_SHAPE_TYPES = {
    "rectangle": 1,
    "oval": 9,
    "line": 10,
    "rounded_rectangle": 5,
}

# ============================================================
# PPTX Chart Types
# ============================================================

PPTX_CHART_TYPES = {
    "bar": 2, "line": 4, "pie": 5, "column": 1,
    "doughnut": -1, "scatter": -1, "bubble": -1,
}

# ============================================================
# XLSX Financial Color Standards (from xlsx skill)
# ============================================================

XLSX_FINANCIAL_COLORS = {
    "hardcoded_input": {"color": "0000FF", "description": "Blue: hardcoded values users may change", "code": "blue"},
    "formula": {"color": "000000", "description": "Black: ALL formulas and calculations", "code": "black"},
    "internal_link": {"color": "008000", "description": "Green: links pulling from same workbook", "code": "green"},
    "external_link": {"color": "FF0000", "description": "Red: external links to other files", "code": "red"},
    "key_assumption_bg": {"color": "FFFF00", "description": "Yellow background: key assumptions needing attention", "code": "yellow_bg"},
}

# ============================================================
# XLSX Number Format Standards (from xlsx skill)
# ============================================================

XLSX_NUMBER_FORMATS = {
    "currency": '$#,##0;($#,##0);-',
    "percentage": '0.0%',
    "multiple": '0.0x',
    "year": '@',  # Text format
    "integer": '#,##0',
    "decimal_2": '#,##0.00',
}

# ============================================================
# XLSX Cell Types (for auto-detection)
# ============================================================

XLSX_CELL_ROLE = {
    "blue": "hardcoded input — should be moved to assumption cell",
    "black": "formula / calculation — correct",
    "green": "internal reference — correct",
    "red": "external reference — verify link validity",
}

# ============================================================
# DOCX DXA Reference (from docx skill)
# ============================================================

DOCX_DXA = {
    "inch_to_dxa": 1440,
    "cm_to_dxa": 567,
    "a4_width": 11906, "a4_height": 16838,
    "a3_width": 16838, "a3_height": 23811,
    "a5_width": 8392, "a5_height": 11906,
    "letter_width": 12240, "letter_height": 15840,
    "legal_width": 12240, "legal_height": 20160,
    "b5_width": 10126, "b5_height": 14388,
    # US Letter with 1" margins: 12240 - 1440*2 = 9360
    "us_letter_content_width_1inch": 9360,
    "a4_content_width_1inch_approx": 9026,
}

# ============================================================
# DOCX Table Rules (from docx skill)
# ============================================================

DOCX_TABLE_RULES = {
    "shading_type": "CLEAR",  # Never SOLID — causes black backgrounds
    "width_type": "DXA",      # Never PERCENTAGE — breaks in Google Docs
    "dual_widths": True,      # Both columnWidths array AND cell width must match
    "width_sum_rule": "table width = sum of columnWidths",
    "cell_padding": {"top": 80, "bottom": 80, "left": 120, "right": 120},  # DXA
    "no_tables_as_dividers": True,  # Use paragraph border instead
}

# ============================================================
# DOCX Page Setup Quick Reference
# ============================================================

DOCX_PAGE_MARGINS = {
    "gb_t9704": {"top": 105, "bottom": 99, "left": 79, "right": 74},  # GB/T 9704 公文
    "thesis_default": {"top": 85, "bottom": 71, "left": 85, "right": 71},
    "standard_1inch": {"top": 72, "bottom": 72, "left": 72, "right": 72},  # 72pt = 1inch = 2.54cm
    "narrow": {"top": 36, "bottom": 36, "left": 36, "right": 36},
}

# ============================================================
# WPS Layout Constants
# ============================================================

WPS_PAGE_ORIENTATION = {0: "portrait", 1: "landscape"}
WPS_TEXT_COLUMNS = {"count": 1, "space": 720, "separate": False}  # default

# ============================================================
# Smart Quotes Map (from docx + pptx skills)
# ============================================================

SMART_QUOTES = {
    "\u2018": "&#x2018;",  # left single '
    "\u2019": "&#x2019;",  # right single '
    "\u201C": "&#x201C;",  # left double "
    "\u201D": "&#x201D;",  # right double "
}

# ============================================================
# Font Reference — Common Chinese Document Fonts
# ============================================================

CN_FONTS = {
    "黑体": {"name": "黑体", "en_name": "SimHei"},
    "宋体": {"name": "宋体", "en_name": "SimSun"},
    "仿宋": {"name": "仿宋", "en_name": "FangSong"},
    "楷体": {"name": "楷体", "en_name": "KaiTi"},
    "微软雅黑": {"name": "微软雅黑", "en_name": "Microsoft YaHei"},
    "方正小标宋简体": {"name": "方正小标宋简体", "en_name": "FZXiaoBiaoSong-B05S"},
}

CN_FONT_SIZE_MAP = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
}

# ============================================================
# Helper: get palette by keyword
# ============================================================

def suggest_palette(topic: str = "") -> dict:
    """Suggest a color palette based on topic keywords. Returns palette dict + rationale."""
    topic_lower = topic.lower()
    mappings = {
        "金融": "midnight_executive", "finance": "midnight_executive",
        "科技": "ocean_gradient", "tech": "ocean_gradient",
        "自然": "forest_moss", "nature": "forest_moss", "环境": "forest_moss",
        "创意": "coral_energy", "creative": "coral_energy",
        "极简": "charcoal_minimal", "minimal": "charcoal_minimal",
        "医疗": "teal_trust", "health": "teal_trust", "medical": "teal_trust",
        "教育": "sage_calm", "education": "sage_calm",
        "时尚": "berry_cream", "fashion": "berry_cream",
        "工业": "warm_terracotta", "industrial": "warm_terracotta",
    }
    for kw, palette_key in mappings.items():
        if kw in topic_lower:
            pal = PPTX_PALETTES.get(palette_key)
            return {"key": palette_key, "palette": pal, "rationale": f"匹配关键词: {kw}"}
    # default
    pal = PPTX_PALETTES["ocean_gradient"]
    return {"key": "ocean_gradient", "palette": pal, "rationale": "通用推荐"}


# ============================================================
# Helper: suggest typography for a style
# ============================================================

def suggest_typography(style: str = "professional") -> dict:
    """Suggest a font pairing based on style."""
    mapping = {
        "classic": 0, "bold": 1, "modern": 2, "formal": 3,
        "friendly": 4, "strong": 5, "literary": 6, "technical": 7,
    }
    idx = mapping.get(style, 2)
    return PPTX_TYPOGRAPHY[idx]
