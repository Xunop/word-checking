FORMAT_STANDARDS = {
    "论文正文": {  # 对应中文环境下的 "正文" 样式
        "font_name_east_asia": "宋体",  # 示例：期望中文字体
        "font_name_latin": "Times New Roman",  # 示例：期望西文字体
        "font_size": 12,  # 磅
        "bold": False,
        "italic": False,
        "alignment": "两端对齐",  # WD_ALIGN_PARAGRAPH.JUSTIFY
        "line_spacing_rule": "多倍行距",  # WD_LINE_SPACING.MULTIPLE
        "line_spacing": 1.25,  # 如果规则是多倍行距，此值为倍数；如果是固定值，此值为磅
        "first_line_indent_cm": 0.74,  # 首行缩进 (约等于2个中文字符，按宋体小四号算) - 需要转换为Pt或Inches比较
        "space_before_pt": 0,  # 段前间距 (磅)
        "space_after_pt": 0,  # 段后间距 (磅)
    },
    "Heading 1": {  # 对应 "标题 1"
        "font_name_east_asia": "黑体",
        "font_name_latin": "Arial",
        "font_size": 22,  # 磅
        "bold": True,
        "italic": False,
        "alignment": "左对齐",  # WD_ALIGN_PARAGRAPH.LEFT
        "line_spacing_rule": "单倍行距",
        "line_spacing": 1.0,
        "first_line_indent_cm": 0,
        "space_before_pt": 12,
        "space_after_pt": 6,
    },
    # "论文正文": {  # 用户提供的样式名
    #     "font_name_east_asia": "宋体",
    #     "font_name_latin": "Times New Roman",
    #     "font_size": 12,
    #     "bold": False,
    #     "italic": False,
    #     "alignment": "两端对齐",
    #     "line_spacing_rule": "固定值",  # 示例，假设为固定值
    #     "line_spacing": 15,  # 固定值15磅 (对应XML <w:spacing w:line="300" .../>) 300twips = 15pt
    #     "first_line_indent_cm": 0.74,  # 约2字符
    #     "space_before_pt": 0,
    #     "space_after_pt": 0,
    # },
    # ... 可以添加更多样式标准 ...
}
