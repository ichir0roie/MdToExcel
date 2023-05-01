from module.markdown_to_excel import MarkdownToExcel

mte = MarkdownToExcel()

text = [
    "| test | test2 | test3 |",
    "|-|  - | - |",
    "| a  | b | c |",
]

for i in text:
    d = mte._get_markdown_line(i)
    print(d)


raise
