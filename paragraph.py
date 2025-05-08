from docx.oxml.ns import qn # qn 用于生成带命名空间的标签名
from docx.shared import Twips
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

# _DOC_DEFAULTS_CACHE = {} # 用于缓存每个文档的默认设置

def get_effective_paragraph_property(paragraph, property_name):
    """
    获取段落指定格式属性的有效值，模拟 Word 的样式解析逻辑。
    property_name 必须是 ParagraphFormat 对象的有效属性名 (字符串)。
    """
    doc = paragraph.part.document # 获取 Paragraph 所在的 Document 对象

    # doc_defaults = get_document_default_pPr(doc)

    # 1. 检查直接格式化
    #    getattr 用于通过字符串名称访问对象的属性
    direct_value = getattr(paragraph.paragraph_format, property_name, None)
    if direct_value is not None:
        return direct_value

    # 2. 检查段落的显式样式
    current_style = paragraph.style 
    if current_style and current_style.type == 1: # WD_STYLE_TYPE.PARAGRAPH
        style_value = getattr(current_style.paragraph_format, property_name, None)
        if style_value is not None:
            return style_value

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
    #    这里简单返回 None，调用者可根据需要处理。
    if property_name in ['keep_together', 'keep_with_next', 'page_break_before', 'widow_control']:
        return False # Word 的常见默认值
    if property_name == 'alignment':
        return WD_PARAGRAPH_ALIGNMENT.LEFT # Word 的常见默认值
    if property_name == 'line_spacing_rule':
        # 如果 line_spacing 也是 None，则通常是 SINGLE
        # 否则，根据 line_spacing 的类型 (float vs Length) 决定
        # 此处简化处理
        return WD_LINE_SPACING.SINGLE
    if property_name == 'line_spacing':
        # 对应 SINGLE rule，line_spacing 通常是 1.0 (float) 或等效 Length
        return 1.0 # 对应 WD_LINE_SPACING.SINGLE 的常见值

    # 对于 Length 类型的属性，如缩进和间距，Word 通常默认为 0
    if property_name in ['first_line_indent', 'left_indent', 'right_indent', 'space_before', 'space_after']:
        return Twips(0) # 0 长度

    return None # 未在文档中找到定义

def get_effective_first_line_indent(paragraph):
    return get_effective_paragraph_property(paragraph, 'first_line_indent')

def get_effective_alignment(paragraph):
    return get_effective_paragraph_property(paragraph, 'alignment')

def get_effective_line_spacing_rule(paragraph):
    return get_effective_paragraph_property(paragraph, 'line_spacing_rule')

def get_effective_line_spacing(paragraph):
    return get_effective_paragraph_property(paragraph, 'line_spacing')
