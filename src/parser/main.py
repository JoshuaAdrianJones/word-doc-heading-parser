# read docx file into headers
from dataclasses import dataclass
from pathlib import Path
from typing import List
from docx import Document
from openpyxl import Workbook


@dataclass
class WordDocument:
    path: Path

    @classmethod
    def make(cls, path: Path) -> "WordDocument":
        return WordDocument(path=path)


@dataclass
class Header:
    level: int
    text: str
    children: List["Header"] | None = None

    def add_child(self, child):
        if self.children is None:
            self.children = []
        self.children.append(child)


if __name__ == "__main__":
    input_file = Path("tests", "fixtures", "parser_test.docx")
    project_document = WordDocument.make(input_file)
    word_document = Document(project_document.path)
    content = word_document.paragraphs
    headers = [para for para in content if para.style.name.startswith("Heading")]
    list_of_headers = []
    for i, header in enumerate(reversed(headers)):
        header_type = header.style.name
        match header_type:
            case "Heading 1":
                level = 1
            case "Heading 2":
                level = 2
            case "Heading 3":
                level = 3
            case _:
                level = 0
        list_of_headers.append(Header(level=level, text=header.text))

    def build_header_hierarchy(headers: List[Header]) -> Header:
        root_header = Header(0, "ROOT")
        level_1_header = None
        level_2_header = None
        for header in reversed(headers):
            if header.level == 1:
                root_header.add_child(header)
                level_1_header = header
                level_2_header = None
            elif header.level == 2:
                level_1_header.add_child(header)
                level_2_header = header
            elif header.level == 3:
                if level_2_header is None:
                    level_2_header = Header(2, "")
                    level_1_header.add_child(level_2_header)
                level_2_header.add_child(header)
        return root_header

    root_header = build_header_hierarchy(list_of_headers)

    def flatten_root(root_header: Header) -> list[str]:
        headers_list = []
        for child in root_header.children:
            print(child.text)
            headers_list.append((child.level, child.text))
            if child.children:
                for childchild in child.children:
                    print(childchild.text)
                    headers_list.append((childchild.level, childchild.text))
                    if childchild.children:
                        for childchildchild in childchild.children:
                            print(childchildchild.text)
                            headers_list.append(
                                (childchildchild.level, childchildchild.text)
                            )
        return headers_list

    wb = Workbook()
    ws = wb.active
    ws.append(["Level", "Heading"])
    for head in flatten_root(root_header=root_header):
        ws.append([head[0], head[1]])
    wb.save("output.xlsx")
