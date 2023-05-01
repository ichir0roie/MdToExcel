from module.markdown_to_excel import MarkdownToExcel

import sys

args = sys.argv[1:]

if not 1 < len(args) < 3:
    raise ValueError("引数が異常")


path_input = None
path_output = None

if len(args) > 1:
    path_input = args[2]
if len(args) > 2:
    path_output = args[3]
