from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE

# 辅助函数：从 lxml 元素中获取字体属性的实际值


def _get_font_property_from_xml_rpr(rpr_element, property_name):
    """
    从 <w:rPr> lxml 元素中提取特定的字体属性。
    rpr_element: CT_RPr 类型的 lxml 元素。
    property_name: 'size', 'name', 'bold', 'italic', etc.
    """
    if rpr_element is None:
        return None

    if property_name == 'size':
        sz_element = rpr_element.find('.//w:sz', namespaces=rpr_element.nsmap)
        if sz_element is not None and sz_element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'):
            # print(sz_element.xml)
            # 值以半磅为单位
            return int(sz_element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) / 2
        sz_cs_element = rpr_element.find(
            './/w:szCs', namespaces=rpr_element.nsmap)  # 针对复杂文种的大小
        if sz_cs_element is not None and sz_cs_element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'):
            return int(sz_cs_element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) / 2
        return None
    elif property_name == 'name':
        fonts_element = rpr_element.find(
            './/w:rFonts', namespaces=rpr_element.nsmap)
        # print(fonts_element.xml)
        if fonts_element is not None:
            # 优先顺序：ascii, hAnsi, eastAsia, cs
            # 实际应用中可能需要更复杂的逻辑来根据文本内容确定使用哪个字体
            name = fonts_element.get(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')
            if name:
                return name
            name = fonts_element.get(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi')
            if name:
                return name
            name = fonts_element.get(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
            if name:
                return name
            name = fonts_element.get(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs')
            if name:
                return name
        return None
    elif property_name == 'bold':
        b_element = rpr_element.find('.//w:b', namespaces=rpr_element.nsmap)
        if b_element is not None:
            val = b_element.get(
                # 默认为 true
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
            return not (val == 'false' or val == '0')  # true, 1, on 都是 True
        return None  # None 表示继承，False 表示显式关闭
    elif property_name == 'italic':
        # print(rpr_element.xml)
        i_element = rpr_element.find('.//w:i', namespaces=rpr_element.nsmap)
        if i_element is not None:
            val = i_element.get(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
            return not (val == 'false' or val == '0')
        return None
    # 可以为其他属性（如 color, underline）添加更多逻辑
    return None

# 辅助函数：递归获取样式层级中的属性


def _get_style_hierarchy_property(style_obj, property_name, doc_styles_element, visited_styles=None):
    """
    递归地从样式对象及其基础样式链中获取字体属性。
    style_obj: CharacterStyle 或 ParagraphStyle 对象。
    property_name: 'size', 'name', 'bold', 'italic'.
    doc_styles_element: document.styles.element，用于最终查询 w:docDefaults。
    visited_styles: 用于防止在样式循环引用时无限递归。
    """
    if style_obj is None:
        return None

    if visited_styles is None:
        visited_styles = set()

    # 防止循环引用导致的无限递归
    if style_obj.style_id in visited_styles:
        return None
    visited_styles.add(style_obj.style_id)

    # 1. 检查当前样式对象的.font 属性
    font_prop_val = getattr(style_obj.font, property_name, None)
    if font_prop_val is not None:
        # 对于布尔型属性，None 表示继承，True/False 表示显式设置
        if property_name in ['bold', 'italic']:
            return font_prop_val  # 直接返回 True, False, 或 None
        # 对于大小和名称，如果不是 None，则直接使用
        if property_name in ['size', 'name'] and font_prop_val is not None:
            return font_prop_val

    # 2. 如果当前样式未定义，且存在基础样式，则检查基础样式
    if style_obj.base_style is not None:
        # 创建一个新的 visited_styles 集合副本进行递归调用
        # 以免影响同级或上级递归栈中的状态
        base_style_val = _get_style_hierarchy_property(
            style_obj.base_style, property_name, doc_styles_element, set(
                visited_styles)
        )
        if base_style_val is not None:
            return base_style_val

    # 如果遍历完基础样式链仍为 None，则不在此函数中查询 w:docDefaults
    # w:docDefaults 的查询将在主函数中作为最后手段
    return None


# 辅助函数：从 w:docDefaults 获取文档级默认字体属性
def _get_doc_default_property(doc_styles_element, property_name):
    """
    从 styles.xml 中的 w:docDefaults/w:rPrDefault/w:rPr 获取字体属性。
    doc_styles_element: document.styles.element 对象。
    property_name: 'size', 'name', 'bold', 'italic'.
    """
    if doc_styles_element is None:
        return None

    try:
        # XPath 查询 w:rPrDefault 下的 w:rPr 元素
        # 注意：lxml 的 XPath 需要命名空间
        nsmap = doc_styles_element.nsmap
        # 有些文档可能没有 'w' 前缀映射到主命名空间，而是默认命名空间
        # 为了稳健，可以尝试查找正确的命名空间 URI
        wp_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        # 构造 XPath 时，如果 nsmap 中有 'w'，则使用 'w:'; 否则，需要更复杂的处理
        # 一个简化的方法是假设 'w' 存在或使用 lxml 的本地名查找
        # xpath_query = 'w:docDefaults/w:rPrDefault/w:rPr'
        # rpr_default_elements = doc_styles_element.xpath(xpath_query)

        # 更稳健的方式是直接查找元素，不依赖 'w:' 前缀，而是使用完整的命名空间
        doc_defaults = doc_styles_element.find(f"{{{wp_ns}}}docDefaults")
        if doc_defaults is None:
            return None
        rpr_default = doc_defaults.find(f"{{{wp_ns}}}rPrDefault")
        if rpr_default is None:
            return None
        rpr = rpr_default.find(f"{{{wp_ns}}}rPr")
        if rpr is None:
            return None

        return _get_font_property_from_xml_rpr(rpr, property_name)

    except Exception:  # pylint: disable=broad-except
        # 在 XML 解析出错时返回 None
        return None

# 主函数：获取一个 run 的有效字体属性


def get_effective_font_property(paragraph, run, property_name):
    """
    获取一个 run 对象的有效字体属性值。
    property_name可以是 'size', 'name', 'bold', 'italic'。
    """
    # print("1")
    # 1. 检查直接应用于 run 的格式
    direct_value = getattr(run.font, property_name, None)
    if direct_value is not None:
        # 对于布尔型属性，None 表示继承，True/False 表示显式设置
        if property_name in ['bold', 'italic']:
            return direct_value
        # 对于大小和名称，如果不是 None，则直接使用
        if property_name == 'name' and direct_value is not None:
            return direct_value
        if property_name == 'size' and direct_value is not None:
            return direct_value.pt

    # 获取 document.styles.element 以备后用
    doc_styles_element = run.part.document.styles.element

    # print("2")
    # 2. 检查应用于 run 的字符样式 (run.style)
    # run.style 始终返回一个 CharacterStyle 对象（可能是默认字符样式）
    if run.style and run.style.style_id != 'DefaultParagraphFont':  # 避免对已知几乎为空的默认样式进行不必要的递归
        # (注意: 'DefaultParagraphFont' 检查可能需要更细致，因为它也可能被用户修改或基于其他样式)
        # 但通常它本身不定义具体字体，而是继承。
        # 为简化，这里可以先尝试解析其属性，如果它有显式定义。
        char_style_value = _get_style_hierarchy_property(
            run.style, property_name, doc_styles_element, visited_styles=set())
        if char_style_value is not None:
            if property_name == 'size':
                return char_style_value.pt
            return char_style_value

    # print("3")
    # 3. 检查段落样式 (paragraph.style)
    if paragraph.style:
        # 传递一个新的 visited_styles 集合
        para_style_value = _get_style_hierarchy_property(
            paragraph.style, property_name, doc_styles_element, visited_styles=set())
        if para_style_value is not None:
            if property_name == 'size':
                return para_style_value.pt
            return para_style_value

    # print("4")
    # 4. 如果以上均未找到，查询文档级默认设置 (w:docDefaults)
    doc_default_value = _get_doc_default_property(
        doc_styles_element, property_name)
    if doc_default_value is not None:
        return doc_default_value

    # print("5")
    # 5. 如果连 w:docDefaults 都没有，则返回 None (或一个应用程序级别的假定默认值)
    # 例如，Word 的普遍默认字体大小可能是 11pt，字体可能是 Calibri。
    # 但 python-docx 无法知晓 MS Word 应用程序的内部默认值。
    if property_name == 'size':
        # 可以考虑返回一个标准的默认值，如 Word 中常见的 11pt，但这超出了文档本身的信息
        # return Pt(11) # 示例：如果希望在完全未指定时有一个回退值
        pass
    elif property_name == 'name':
        # return "Calibri" # 示例
        pass
    elif property_name == 'italic':
        return False
    elif property_name == 'bold':
        return False

    return None  # 表示在文档内无法确定
