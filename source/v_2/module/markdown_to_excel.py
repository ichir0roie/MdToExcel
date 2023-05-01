
import os
import csv


class MarkdownToExcel:
    def __init__(self) -> None:
        pass

    table_header_characters = set(["|", "-", " "])
    header_characters = set(["#"])

    def run(self, path_input: str, path_output: str = None):
        text = self.read_text_data(path_input)

    def read_file(self, path: str) -> str:
        if os.path.splitext(os.path.basename(path))[-1] != ".md":
            raise Exception("File is not markdown file.")
        with open(path, "r", encoding="utf-8") as f:
            text = f.read()
        return text

    def read_text_data(self, text: str) -> list[list[str]]:
        data = [
            self._get_markdown_line(line)
            for line in text.split("\n")
            if line is not None
        ]

        return data

    def markdown_data_to_tsv_data(self, data: list[list[str]]) -> list[list[str]]:

        tsv_data = []

        indent = 0
        indent_counts = [0 for _ in range(10)]

        for line in data:
            header_line = False
            if not line:
                tsv_data.append([])
                continue
            if set(line[0]) == self.header_characters:
                indent = len(line[0])-1

                indent_counts[indent] += 1
                for i in range(indent+1, len(indent_counts)):
                    indent_counts[i] = 0

                header_line = True
                indent_text = ".".join(
                    [
                        str(txt_indent)
                        for txt_indent in indent_counts[:indent+1]
                    ]
                )+"."
                line = [indent_text]+line[1:]

            add_indent = indent+1 if not header_line else indent

            tsv_data.append(
                ["" for _ in range(add_indent)]+line
            )
        return tsv_data

    def _get_markdown_line(self, line: str) -> list[str]:
        characters = set(list(line))
        if characters == self.table_header_characters:
            return None
        if line == "":
            return []

        if line[0] == "|" and line[-1] == "|":
            return self._get_table_line(line)
        else:
            return self._get_normal_text(line)

    def _get_normal_text(self, line: str) -> list[str]:
        t = line.replace("  ", "\t")
        t = t.replace("\t\t", "\t")
        t = t.replace("\t ", "\t")
        t = t.replace("# ", "#\t")
        return t.split("\t")

    def _get_table_line(self, line: str) -> list[str]:

        t = line.replace(" |", "|")
        t = t.replace("| ", "|")
        return t.split("|")[1:-1]

    def write_list_to_tsv(self, path: str = None, data: list[list[str]] = []) -> None:
        if os.path.splitext(os.path.basename(path))[-1] != ".tsv":
            raise Exception("output is not tsv file.")
        with open(path, "w", encoding="utf-8", newline="") as f:
            write = csv.writer(f, delimiter="\t")
            write.writerows(data)
