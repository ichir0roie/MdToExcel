from module.markdown_to_excel import MarkdownToExcel

mte = MarkdownToExcel()

indent_text = """
| test | test2 | test3 |
|-|  - | - |
| a  | b | c |
test text this is.
    tab in  tab
    tab is tab
        second tab is this
        

# header 1        

| test | test2 | test3 |
|-|  - | - |
| a  | b | c |

## header 2

test text this is.
    tab in  tab
    tab is tab
        second tab is this
   
# header 1   

a

## header 2

b

## header 2

b

## header 2

b

### header 3

b

### header 3

c
b

## header 2

b

### header 3

c


"""

data = mte.read_text_data(indent_text)
res = mte.markdown_data_to_tsv_data(data)

mte.write_list_to_tsv("out.tsv", res)

raise
