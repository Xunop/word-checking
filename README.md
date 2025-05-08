# word-checking

折腾了一天这个脚本，我选择手动排查我的论文格式😢。

实现的最终效果很模糊，需要花些时间去了解 [OOXML 的规范](http://officeopenxml.com/anatomyofOOXML.php) 和 [python-docx](https://python-docx.readthedocs.io/en/latest/index.html)，实在没这个时间。~~有这时间我都能找到好多格式问题了~~

脚本只能检查 docx 格式，不是 Strict Office Open XML，可以使用在 Microsoft Word 中使用“另存为”转换成 docx 格式。

## How to use

修改 `rules.py` 中的规则集格式，复制一份文档到代码所在目录，修改 `checking.py` 中的文件名：

```py
    doc_file_path = (
        "test.docx"
    )
```

```bash
pip install -r requirements.txt
python checking.py
```

## FAQ

1. 样式可能与 Microsoft Word 中显示的不同，比如 “正文” 会被检测成 “Normarl”
