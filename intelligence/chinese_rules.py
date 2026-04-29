# -*- coding: utf-8 -*-
# All measurements in points (1pt ≈ 0.3528mm, 1mm ≈ 2.835pt)
# Font sizes: 二号=22pt, 三号=16pt, 四号=14pt, 小四=12pt, 五号=10.5pt
CHINESE_FORMATTING = {
    "official": {
        "标题": {
            "font_name": "方正小标宋简体", "font_name_fallback": "宋体",
            "font_size": 22,
            "bold": False, "alignment": "center",
            "line_spacing_rule": "exactly", "line_spacing": 28
        },
        "一级标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 6, "space_after": 3,
        },
        "二级标题": {
            "font_name": "楷体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 3, "space_after": 3,
        },
        "三级标题": {
            "font_name": "仿宋", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 3,
        },
        "正文": {
            "font_name": "仿宋", "font_size": 16,
            "first_line_indent_chars": 2,
            "line_spacing_rule": "exactly", "line_spacing": 28,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 105, "bottom_margin": 99,
            "left_margin": 79, "right_margin": 74,
        }
    },
    "thesis": {
        "章标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_before": 12, "space_after": 12,
        },
        "节标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 12, "space_after": 6,
        },
        "条标题": {
            "font_name": "黑体", "font_size": 12,
            "bold": True, "alignment": "left",
            "outline_level": 3, "space_before": 6, "space_after": 3,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent_chars": 2,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 85, "bottom_margin": 71,
            "left_margin": 85, "right_margin": 71,
        }
    },
    "report": {
        "封面标题": {
            "font_name": "微软雅黑", "font_size": 26,
            "bold": True, "alignment": "center",
            "is_cover": True,
        },
        "一级标题": {
            "font_name": "微软雅黑", "font_size": 18,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "二级标题": {
            "font_name": "微软雅黑", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 6, "space_after": 3,
        },
        "正文": {
            "font_name": "微软雅黑", "font_size": 11,
            "first_line_indent_chars": 2,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 90, "right_margin": 90,
        }
    },
    "resume": {
        "姓名": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "is_cover": True,
        },
        "章节标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "正文": {
            "font_name": "宋体", "font_size": 11,
            "first_line_indent_chars": 0,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 90, "right_margin": 90,
        }
    },
    "contract": {
        "合同标题": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "is_cover": True, "space_after": 12,
        },
        "条款标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 85, "bottom_margin": 85,
            "left_margin": 99, "right_margin": 99,
        }
    },
    "letter": {
        "标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "收件人": {
            "font_name": "仿宋", "font_size": 14,
            "alignment": "left",
        },
        "正文": {
            "font_name": "仿宋", "font_size": 14,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 79, "right_margin": 74,
        }
    },
    "proposal": {
        "封面标题": {
            "font_name": "黑体", "font_size": 26,
            "bold": True, "alignment": "center",
            "is_cover": True, "space_after": 18,
        },
        "一级标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "二级标题": {
            "font_name": "楷体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 6, "space_after": 3,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 90, "right_margin": 90,
        }
    },
    "meeting_minutes": {
        "会议标题": {
            "font_name": "黑体", "font_size": 18,
            "bold": True, "alignment": "center",
            "is_cover": True, "space_after": 12,
        },
        "议题标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "正文": {
            "font_name": "仿宋", "font_size": 14,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 79, "right_margin": 74,
        }
    },
    "press_release": {
        "标题": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "副标题": {
            "font_name": "楷体", "font_size": 14,
            "alignment": "center", "space_after": 6,
        },
        "来源": {
            "font_name": "楷体", "font_size": 12,
            "alignment": "center", "space_after": 6,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 90, "right_margin": 90,
        }
    },
    "manual": {
        "章标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "节标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 6, "space_after": 3,
        },
        "步骤": {
            "font_name": "宋体", "font_size": 11,
            "bold": True, "alignment": "left",
            "outline_level": 3, "space_before": 3, "space_after": 3,
        },
        "正文": {
            "font_name": "宋体", "font_size": 11,
            "first_line_indent_pt": 0,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 57, "bottom_margin": 57,
            "left_margin": 71, "right_margin": 71,
        }
    },
    "exam": {
        "试卷标题": {
            "font_name": "黑体", "font_size": 18,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "大题标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 12, "space_after": 6,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent_pt": 0,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 90, "right_margin": 90,
        }
    },
    "bid": {
        "标书标题": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "is_cover": True, "space_after": 12,
        },
        "一级标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "二级标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 6, "space_after": 3,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 90, "right_margin": 90,
        }
    },
    "notice": {
        "标题": {
            "font_name": "黑体", "font_size": 18,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "主送机关": {
            "font_name": "仿宋", "font_size": 14,
            "alignment": "left", "space_after": 6,
        },
        "正文": {
            "font_name": "仿宋", "font_size": 14,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 79, "right_margin": 74,
        }
    },
    "work_report": {
        "标题": {
            "font_name": "黑体", "font_size": 18,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "一级标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "二级标题": {
            "font_name": "楷体", "font_size": 14, "font_name_fallback": "宋体",
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 6, "space_after": 3,
        },
        "正文": {
            "font_name": "仿宋", "font_size": 14,
            "first_line_indent_pt": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 72, "bottom_margin": 72,
            "left_margin": 79, "right_margin": 74,
        }
    },
}
