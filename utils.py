import tangled_up_in_unicode as unicodedata_tuu
import docx
from docx.oxml.ns import qn
from docx.shared import Pt, Cm  # 用于处理磅和厘米单位
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING  # 用于段落格式


def get_character_script(char):
    try:
        return unicodedata_tuu.script(char)
    except ValueError:
        return "Unknown"


def is_punctuation(char):
    # unicodedata_tuu 没有直接的 is_punctuation，但可以检查类别
    # 'P' 开头的类别通常是标点符号
    return unicodedata_tuu.category(char).startswith('P') if char and not char.isspace() else False


def is_cjk_char(char):
    script = get_character_script(char)
    return script in ["Han", "Hiragana", "Katakana", "Hangul", "Bopomofo"]


def is_latin_char(char):
    return get_character_script(char) == "Latin"


def get_style_rfonts_attr(style, attr_name):
    if style is None or not hasattr(style, "_element") or style._element is None:
        return None
    rpr = style._element.rPr
    if rpr is None:
        return None
    rfonts = rpr.rFonts
    if rfonts is None:
        return None
    return rfonts.get(qn(attr_name))


def get_default_rfonts_attr(document, attr_name):
    try:
        styles_element = document.styles.element
        xpath_query = "./w:docDefaults/w:rPrDefault/w:rPr/w:rFonts"
        rfonts_elements = styles_element.xpath(xpath_query)
        if rfonts_elements:
            return rfonts_elements[0].get(qn(attr_name))
    except Exception:
        pass
    return None


def get_effective_run_fonts(run, paragraph, document):
# def get_effective_run_fonts(run, paragraph):
    effective_fonts = {"ascii": None,
                       "hAnsi": None, "eastAsia": None, "cs": None}
    attr_names = list(effective_fonts.keys())

    run_element = run._element
    rpr = run_element.rPr
    if rpr is not None:
        rfonts = rpr.rFonts
        if rfonts is not None:
            for attr in attr_names:
                if effective_fonts[attr] is None:
                    val = rfonts.get(qn(f"w:{attr}"))
                    if val:
                        effective_fonts[attr] = val

    char_style = run.style
    if char_style and char_style.type == docx.enum.style.WD_STYLE_TYPE.CHARACTER:  # 确保是字符样式
        for attr in attr_names:
            if effective_fonts[attr] is None:
                val = get_style_rfonts_attr(
                    char_style.font, f"w:{attr}")  # 字符样式字体在 style.font
                if val:
                    effective_fonts[attr] = val

    para_style = paragraph.style
    if para_style and para_style.type == docx.enum.style.WD_STYLE_TYPE.PARAGRAPH:  # 确保是段落样式
        for attr in attr_names:
            if effective_fonts[attr] is None:
                # 段落样式字体在其 font 属性
                val = get_style_rfonts_attr(para_style.font, f"w:{attr}")
                if val:
                    effective_fonts[attr] = val

    # 检查文档默认设置
    for attr in attr_names:
        if effective_fonts[attr] is None:
            val = get_default_rfonts_attr(document, f"w:{attr}")
            if val:
                effective_fonts[attr] = val
    return effective_fonts
