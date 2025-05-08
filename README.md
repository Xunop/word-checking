# word-checking

折腾了一天这个脚本，我选择手动排查 word 论文格式👋。最终效果很模糊，需要花些时间去了解 [OOXML 的规范](http://officeopenxml.com/anatomyofOOXML.php) 和 [python-docx](https://python-docx.readthedocs.io/en/latest/index.html)，实在没这个时间。~~有这时间我都能找到好多格式问题了~~

## How to use

修改 `rules.py` 中的规则集格式，复制一份文档到代码所在目录，修改 `checking.py` 中的路径：

```py
    doc_file_path = (
        "test.docx"
    )
```

```bash
pip install -r requirements.txt
python checking.py
```
