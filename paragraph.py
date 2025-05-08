from docx.oxml.ns import qn # qn 用于生成带命名空间的标签名
from docx.shared import Length, Twips
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING

# 简化的 OOXML 到 Python 类型的转换函数示例
# 注意：实际的 python-docx 内部转换更为复杂和健壮

def _parse_ooxml_boolean(pPr_element, tag_name):
    """解析 OOXML 的 on/off 元素为布尔值。"""
    el = pPr_element.find(qn(tag_name))
    if el is not None:
        val = el.get(qn('w:val'))
        if val is None or val in ('1', 'true', 'on'):
            return True
        if val in ('0', 'false', 'off'):
            return False
    return None # 未定义则返回 None，以便继承

def _parse_ooxml_length(pPr_element, parent_tag, attr_name):
    """解析 OOXML 长度属性 (twips) 为 Length 对象 (EMU)。"""
    parent_el = pPr_element.find(qn(parent_tag))
    if parent_el is not None:
        val_str = parent_el.get(qn(attr_name))
        if val_str:
            return Twips(int(val_str))
    return None

def _parse_ooxml_alignment(pPr_element):
    """解析 OOXML 对齐方式为 WD_PARAGRAPH_ALIGNMENT 枚举。"""
    jc_el = pPr_element.find(qn('w:jc'))
    if jc_el is not None:
        val_str = jc_el.get(qn('w:val'))
        # 这是一个简化的映射，实际 python-docx 枚举处理更完善
        mapping = {
            "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "both": WD_PARAGRAPH_ALIGNMENT.JUSTIFY, # 'both' 通常对应 JUSTIFY
            "distribute": WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE,
            # 添加其他必要的映射
        }
        return mapping.get(val_str)
    return None

def _parse_ooxml_line_spacing_rule(pPr_element):
    """解析 OOXML 行距规则。"""
    spacing_el = pPr_element.find(qn('w:spacing'))
    if spacing_el is not None:
        rule_str = spacing_el.get(qn('w:lineRule'))
        # 简化映射，实际转换见 parfmt.py
        if rule_str == "exact": return WD_LINE_SPACING.EXACTLY
        if rule_str == "atLeast": return WD_LINE_SPACING.AT_LEAST
        if rule_str == "auto" or rule_str is None: # 'auto' 或无 rule 通常表示 MULTIPLE 或 SINGLE
            # 需要结合 w:line 值判断是 SINGLE, ONE_POINT_FIVE, DOUBLE 还是其他 MULTIPLE
            # 此处简化为 MULTIPLE
            line_str = spacing_el.get(qn('w:line'))
            if line_str:
                line_val = int(line_str)
                if rule_str is None and line_val == 240: # Word UI 单倍行距通常是 <w:spacing w:line="240"/> 无 lineRule
                     return WD_LINE_SPACING.SINGLE
                if line_val == 240 and (rule_str == "auto" or rule_str is None): return WD_LINE_SPACING.SINGLE
                if line_val == 360 and (rule_str == "auto" or rule_str is None): return WD_LINE_SPACING.ONE_POINT_FIVE
                if line_val == 480 and (rule_str == "auto" or rule_str is None): return WD_LINE_SPACING.DOUBLE
            return WD_LINE_SPACING.MULTIPLE 
    return None


def _parse_ooxml_line_spacing(pPr_element):
    """解析 OOXML 行距值。"""
    spacing_el = pPr_element.find(qn('w:spacing'))
    if spacing_el is not None:
        line_str = spacing_el.get(qn('w:line'))
        if line_str:
            line_val = int(line_str)
            rule_str = spacing_el.get(qn('w:lineRule'))
            if rule_str in ("exact", "atLeast"):
                return Twips(line_val)
            if rule_str == "auto" or rule_str is None: # 'auto' 或无 rule
                return float(line_val) / 240.0
    return None


def get_document_default_pPr(document):
    """
    解析并返回文档的默认段落属性 (<w:docDefaults><w:pPrDefault><w:pPr>)。
    返回一个字典，键为 python-docx ParagraphFormat 属性名，值为解析后的 Python 对象。
    """
    doc_defaults_pPr_dict = {}
    if document.styles.element is None: # Should not happen for valid docx
        return doc_defaults_pPr_dict

    # XPath to find the <w:pPr> element under <w:docDefaults>/<w:pPrDefault>
    # 需要完整的命名空间映射
    nspmap = document.styles.element.nsmap
    xpath_query = './w:docDefaults/w:pPrDefault/w:pPr'
    
    # Ensure 'w' prefix is mapped, common in python-docx's lxml usage
    # If not present in doc.styles.element.nsmap, find it from known list
    if 'w' not in nspmap:
        # Attempt to find the main wordprocessingml namespace
        for prefix, uri in nspmap.items():
            if uri == 'http://schemas.openxmlformats.org/wordprocessingml/2006/main':
                nspmap['w'] = uri
                break
        if 'w' not in nspmap: # Still not found, use standard
             nspmap['w'] = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    pPr_elements = document.styles.element.xpath(xpath_query, namespaces=nspmap)

    if not pPr_elements:
        return doc_defaults_pPr_dict # 没有找到 docDefaults pPr

    doc_default_pPr_xml = pPr_elements

    # 解析各个属性
    # 缩进 (Indentation)
    ind_el = doc_default_pPr_xml.find(qn('w:ind'))
    if ind_el is not None:
        val_str = ind_el.get(qn('w:firstLine'))
        if val_str: doc_defaults_pPr_dict['first_line_indent'] = Twips(int(val_str))
        else:
            val_str = ind_el.get(qn('w:hanging'))
            if val_str: doc_defaults_pPr_dict['first_line_indent'] = Twips(-int(val_str))
        
        val_str = ind_el.get(qn('w:left')) # 或 w:start
        if val_str: doc_defaults_pPr_dict['left_indent'] = Twips(int(val_str))
        
        val_str = ind_el.get(qn('w:right')) # 或 w:end
        if val_str: doc_defaults_pPr_dict['right_indent'] = Twips(int(val_str))

    # 对齐 (Alignment)
    doc_defaults_pPr_dict['alignment'] = _parse_ooxml_alignment(doc_default_pPr_xml)

    # 间距 (Spacing)
    spacing_el = doc_default_pPr_xml.find(qn('w:spacing'))
    if spacing_el is not None:
        val_str = spacing_el.get(qn('w:before'))
        if val_str: doc_defaults_pPr_dict['space_before'] = Twips(int(val_str))
        val_str = spacing_el.get(qn('w:after'))
        if val_str: doc_defaults_pPr_dict['space_after'] = Twips(int(val_str))
        
        # 行距规则和行距值 (复杂，依赖于彼此)
        doc_defaults_pPr_dict['line_spacing_rule'] = _parse_ooxml_line_spacing_rule(doc_default_pPr_xml)
        doc_defaults_pPr_dict['line_spacing'] = _parse_ooxml_line_spacing(doc_default_pPr_xml)

    # 布尔型属性
    doc_defaults_pPr_dict['keep_together'] = _parse_ooxml_boolean(doc_default_pPr_xml, 'w:keepLines')
    doc_defaults_pPr_dict['keep_with_next'] = _parse_ooxml_boolean(doc_default_pPr_xml, 'w:keepNext')
    doc_defaults_pPr_dict['page_break_before'] = _parse_ooxml_boolean(doc_default_pPr_xml, 'w:pageBreakBefore')
    doc_defaults_pPr_dict['widow_control'] = _parse_ooxml_boolean(doc_default_pPr_xml, 'w:widowControl')
    
    # 移除值为 None 的条目，以便后续 get() 操作能正确返回 None
    return {k: v for k, v in doc_defaults_pPr_dict.items() if v is not None}

def get_effective_paragraph_property(paragraph, property_name):
    """
    获取段落指定格式属性的有效值，模拟 Word 的样式解析逻辑。
    property_name 必须是 ParagraphFormat 对象的有效属性名 (字符串)。
    """
    doc = paragraph.part.document # 获取 Paragraph 所在的 Document 对象

    # doc_defaults = get_document_default_pPr(doc)

    # print("1")
    # 1. 检查直接格式化
    direct_value = getattr(paragraph.paragraph_format, property_name, None)
    if direct_value is not None:
        return direct_value

    # print("2")
    # 2. 检查段落的显式样式
    current_style = paragraph.style 
    if current_style and current_style.type == 1: # WD_STYLE_TYPE.PARAGRAPH
        style_value = getattr(current_style.paragraph_format, property_name, None)
        if style_value is not None:
            return style_value
        # print("3")
        # 3. 遍历基样式层级
        #    确保 current_style 是 ParagraphStyle 类型，它才有 base_style
        #    CharacterStyle (type 2) 可能作为 base_style，但不含 paragraph_format
        base_style_candidate = current_style.base_style
        while base_style_candidate:
            if base_style_candidate.type == 1: # WD_STYLE_TYPE.PARAGRAPH
                base_style_value = getattr(base_style_candidate.paragraph_format, property_name, None)
                if base_style_value is not None:
                    return base_style_value
            base_style_candidate = base_style_candidate.base_style # 继续向上查找
            
    # 4. 检查文档默认段落设置
    #    doc_defaults 字典的键应与 property_name 匹配
    # default_value = doc_defaults.get(property_name)
    # if default_value is not None:
    #     return default_value
        
    # print("5")
    # 5. 如果所有层级都未定义，python-docx 的 ParagraphFormat 属性本身
    #    在被访问时，如果其底层 XML 不存在对应设置，通常会返回 None 或一个
    #    预设的默认值（如 False for booleans）。
    #    此时，我们返回 None，表示文档中未显式定义，依赖应用程序默认。
    #    或者，可以根据 property_name 返回一个更具体的应用程序级默认值，
    #    但这超出了基于文档内容的解析范围。
    #    例如，对于布尔值，Word 通常默认为 False。
    #    对于对齐，通常是 LEFT。
    #    对于缩进/间距，通常是 0 或等效的无缩进/间距。
    #    对于行距，通常是 SINGLE。
    #    这里简单返回 None
    if property_name in ['keep_together', 'keep_with_next', 'page_break_before', 'widow_control']:
        # return False # Word 的常见默认值
        return None
    if property_name == 'alignment':
        # return WD_PARAGRAPH_ALIGNMENT.LEFT # Word 的常见默认值
        return None
    if property_name == 'line_spacing_rule':
        # 如果 line_spacing 也是 None，则通常是 SINGLE
        # 否则，根据 line_spacing 的类型 (float vs Length) 决定
        # 此处简化处理
        # return WD_LINE_SPACING.SINGLE
        return None
    if property_name == 'line_spacing':
        # 对应 SINGLE rule，line_spacing 通常是 1.0 (float) 或等效 Length
        # return 1.0 # 对应 WD_LINE_SPACING.SINGLE 的常见值
        return None

    # 对于 Length 类型的属性，如缩进和间距，Word 通常默认为 0
    if property_name in ['first_line_indent', 'left_indent', 'right_indent', 'space_before', 'space_after']:
        # return Twips(0) # 0 长度
        return None

    return None # 未在文档中找到定义

def get_effective_first_line_indent(paragraph):
    line_indent = get_effective_paragraph_property(paragraph, 'first_line_indent')
    if isinstance(line_indent, Length):
        return line_indent.pt
    return line_indent

def get_effective_alignment(paragraph):
    return get_effective_paragraph_property(paragraph, 'alignment')

def get_effective_line_spacing_rule(paragraph):
    return get_effective_paragraph_property(paragraph, 'line_spacing_rule')

# def get_effective_line_spacing(paragraph):
#     if get_effective_line_spacing_rule(paragraph) is not None:
#
#     return get_effective_paragraph_property(paragraph, 'line_spacing')

def get_effective_font_size_pt_for_paragraph(paragraph):
    """获取段落的有效字体大小（处理继承），返回磅值"""
    size_pt = None
    current_s = paragraph.style
    while current_s:
        if current_s.font and current_s.font.size is not None:
            size_pt = current_s.font.size.pt
            break
        current_s = current_s.base_style
    if size_pt is None: # 回退到 Normal 样式或硬编码默认值
        try:
            doc = paragraph.part.document
            normal_style = doc.styles['Normal']
            if normal_style.font and normal_style.font.size is not None:
                size_pt = normal_style.font.size.pt
        except (AttributeError, KeyError): pass # 忽略错误，使用硬编码默认值
    return size_pt if size_pt is not None else 11.0

def get_effective_line_spacing(paragraph):
    """
    计算并返回段落的有效行距，统一为磅 (points) 值。
    此函数基于用户的原始代码片段和描述进行了修改。
    它依赖外部辅助函数来解析继承的段落属性。
    """

    line_spacing_value = get_effective_paragraph_property(paragraph, 'line_spacing')
    line_spacing_rule = get_effective_line_spacing_rule(paragraph)
    effective_font_size_pt = get_effective_font_size_pt_for_paragraph(paragraph)

    # 如果无法确定字体大小，则提供一个回退默认值
    if effective_font_size_pt is None:
        # print("警告: 无法确定有效字体大小。行距计算将使用默认值 11pt。")
        effective_font_size_pt = 11.0

    # --- 情况1: line_spacing_value 是一个 Length 对象 (例如 Pt(12), Inches(0.5)) ---
    # 这通常意味着行距是一个固定高度 (对应规则 WD_LINE_SPACING.EXACTLY 或 WD_LINE_SPACING.AT_LEAST)
    if isinstance(line_spacing_value, Length):
        return line_spacing_value.pt

    # --- 情况2: line_spacing_value 是一个浮点数 (例如 1.0, 1.5, 2.75) ---
    # 这通常意味着行距是行高的倍数 (对应规则 WD_LINE_SPACING.SINGLE, 
    # WD_LINE_SPACING.ONE_POINT_FIVE, WD_LINE_SPACING.DOUBLE, WD_LINE_SPACING.MULTIPLE)
    # 描述中提到："A float value, e.g. 2.0 or 1.75, indicates spacing is applied in multiples of line heights."
    if isinstance(line_spacing_value, float):
        # 该浮点数即为倍数
        return line_spacing_value * effective_font_size_pt

    # --- 情况3: line_spacing_value 是 None ---
    # 这意味着行距是从样式层级继承的，并且 `get_effective_paragraph_property`
    # 未能将其解析为一个具体的浮点数或 Length 对象，或者它确实是文档的默认设置。
    # 此时，我们需要依赖 `line_spacing_rule`，或者在规则也是 None 的情况下，
    # 假定为文档的默认行距（通常是“单倍行距”）。
    if line_spacing_value is None:
        # Word 中，如果未指定行距，通常默认为“单倍行距”。
        # “单倍行距”的倍数可能是 1.0，或者在较新版本的 Word 中是约 1.08 或 1.15。
        # 此处使用 1.08 作为“现代单倍行距”的占位符，理想情况下应从文档的实际默认值获取。
        default_multiplier = 1.08

        if line_spacing_rule == WD_LINE_SPACING.SINGLE:
            # 如果规则是 SINGLE，但 `line_spacing_value` 为 None，
            # 这表明使用的是默认的单倍行距倍数。
            # python-docx 对于 `line_spacing` 属性，在规则为 SINGLE 时，
            # 通常其值会是 1.0（如果明确设置了单倍行距）。
            # 如果 `get_effective_paragraph_property` 返回了 None，则我们应用一个标准值。
            default_multiplier = 1.0 # 或者 1.08 / 1.15，取决于您希望如何解释“默认单倍”
        elif line_spacing_rule == WD_LINE_SPACING.ONE_POINT_FIVE:
            default_multiplier = 1.5
        elif line_spacing_rule == WD_LINE_SPACING.DOUBLE:
            default_multiplier = 2.0
        elif line_spacing_rule == WD_LINE_SPACING.MULTIPLE:
            # 规则是 MULTIPLE，但 `line_spacing_value` 为 None。这表示倍数未定义。
            # 通常意味着倍数值是从父级继承但未能解析，或者应采用类似单倍行距的默认值。
            # print("警告: 行距规则为 MULTIPLE 但 line_spacing_value 为 None。假定为默认倍数。")
            default_multiplier = 1.08 # 或者 1.0
        elif line_spacing_rule == WD_LINE_SPACING.EXACTLY or \
             line_spacing_rule == WD_LINE_SPACING.AT_LEAST:
            # 规则指示固定高度，但 `line_spacing_value` 为 None。这是一个不一致的状态。
            # 此时应期望得到一个 Length 对象。
            # print("警告: 行距规则为 EXACTLY/AT_LEAST 但 line_spacing_value 为 None。回退到默认行为。")
            # 回退到基于字体大小的默认倍数行距。
            default_multiplier = 1.08
        # 如果 line_spacing_rule 也为 None (例如，段落完全没有指定行距信息)，
        # 将使用上面设定的 `default_multiplier` (例如 1.08)。
        
        return default_multiplier * effective_font_size_pt
        
    # --- 回退处理 ---
    # 如果 `line_spacing_value` 是其他未预期的类型（例如，一个整数但不是 Length 子类），
    # 这表明 `get_effective_paragraph_property` 的返回值不符合描述中的
    # “float or Length value or None”。
    # 在这种情况下，默认返回基于字体大小的“现代单倍行距”。
    print(f"警告: 未处理的 line_spacing_value 类型 ('{type(line_spacing_value)}') 或场景。回退到单倍行距。")
    return 1.08 * effective_font_size_pt # 使用占位符“现代单倍行距”倍数

