# -*- coding: utf-8 -*-
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
            "outline_level": 1,
        },
        "二级标题": {
            "font_name": "楷体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 2,
        },
        "正文": {
            "font_name": "仿宋", "font_size": 16,
            "first_line_indent": 2,
            "line_spacing_rule": "exactly", "line_spacing": 28,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 37, "bottom_margin": 35,
            "left_margin": 28, "right_margin": 26,
        }
    },
    "thesis": {
        "章标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_before": 12, "space_after": 6,
        },
        "节标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2, "space_before": 6, "space_after": 3,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent": 2,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 31.7, "right_margin": 31.7,
        }
    },
    "report": {
        "封面标题": {
            "font_name": "微软雅黑", "font_size": 26,
            "bold": True, "alignment": "center",
            "outline_level": 1,
        },
        "一级标题": {
            "font_name": "微软雅黑", "font_size": 18,
            "bold": True, "alignment": "left",
            "outline_level": 1,
        },
        "二级标题": {
            "font_name": "微软雅黑", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2,
        },
        "正文": {
            "font_name": "微软雅黑", "font_size": 11,
            "first_line_indent": 2,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 31.7, "right_margin": 31.7,
        }
    },
    "resume": {
        "姓名": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "outline_level": 1,
        },
        "章节标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 1,
            "space_before": 12, "space_after": 6,
        },
        "正文": {
            "font_name": "宋体", "font_size": 11,
            "first_line_indent": 0,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 31.7, "right_margin": 31.7,
        }
    },
    "contract": {
        "合同标题": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "条款标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 6, "space_after": 3,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 30, "bottom_margin": 30,
            "left_margin": 35, "right_margin": 35,
        }
    },
    "letter": {
        "标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "center",
            "outline_level": 1,
        },
        "收件人": {
            "font_name": "仿宋", "font_size": 14,
            "alignment": "left",
        },
        "正文": {
            "font_name": "仿宋", "font_size": 14,
            "first_line_indent": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 28, "right_margin": 26,
        }
    },
    "proposal": {
        "封面标题": {
            "font_name": "黑体", "font_size": 26,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 18,
        },
        "一级标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12,
        },
        "二级标题": {
            "font_name": "楷体", "font_size": 15,
            "bold": True, "alignment": "left",
            "outline_level": 2,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 31.7, "right_margin": 31.7,
        }
    },
    "meeting_minutes": {
        "会议标题": {
            "font_name": "黑体", "font_size": 18,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "议题标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 1,
        },
        "正文": {
            "font_name": "仿宋", "font_size": 14,
            "first_line_indent": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 28, "right_margin": 26,
        }
    },
    "press_release": {
        "标题": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "outline_level": 1,
        },
        "副标题": {
            "font_name": "楷体", "font_size": 14,
            "alignment": "center",
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 31.7, "right_margin": 31.7,
        }
    },
    "manual": {
        "章标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12,
        },
        "节标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2,
        },
        "步骤": {
            "font_name": "宋体", "font_size": 11,
            "bold": True, "alignment": "left",
            "outline_level": 3,
        },
        "正文": {
            "font_name": "宋体", "font_size": 11,
            "first_line_indent": 0,
            "line_spacing_rule": "multiple", "line_spacing": 1.3,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 20, "bottom_margin": 20,
            "left_margin": 25, "right_margin": 25,
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
            "outline_level": 2, "space_before": 12,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent": 0,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "left",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 31.7, "right_margin": 31.7,
        }
    },
    "bid": {
        "标书标题": {
            "font_name": "黑体", "font_size": 22,
            "bold": True, "alignment": "center",
            "outline_level": 1, "space_after": 12,
        },
        "一级标题": {
            "font_name": "黑体", "font_size": 16,
            "bold": True, "alignment": "left",
            "outline_level": 1, "space_before": 12,
        },
        "二级标题": {
            "font_name": "黑体", "font_size": 14,
            "bold": True, "alignment": "left",
            "outline_level": 2,
        },
        "正文": {
            "font_name": "宋体", "font_size": 12,
            "first_line_indent": 28,
            "line_spacing_rule": "multiple", "line_spacing": 1.5,
            "alignment": "justify",
        },
        "page": {
            "paper": "A4", "top_margin": 25.4, "bottom_margin": 25.4,
            "left_margin": 31.7, "right_margin": 31.7,
        }
    },
}
