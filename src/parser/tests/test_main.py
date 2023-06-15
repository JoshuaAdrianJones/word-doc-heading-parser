from pathlib import Path

from ..main import Header, build_header_hierarchy, load_headers_from_file


def test_add_child():
    header1 = Header(1, "Header 1")
    header2 = Header(2, "Header 2")
    header3 = Header(3, "Header 3")

    # Add child to header1
    header1.add_child(header2)
    assert len(header1.children) == 1
    assert header1.children[0] == header2

    # Add child to header2
    header2.add_child(header3)
    assert len(header2.children) == 1
    assert header2.children[0] == header3

    # Test adding child to header with no children
    header4 = Header(1, "Header 4")
    header4.add_child(header3)
    assert len(header4.children) == 1
    assert header4.children[0] == header3


def test_can_compare_documents() -> None:
    list_of_headers_1 = load_headers_from_file(
        Path("tests", "fixtures", "parser_test.docx")
    )
    list_of_headers_2 = load_headers_from_file(
        Path("tests", "fixtures", "parser_test_match.docx")
    )
    root_header_1 = build_header_hierarchy(list_of_headers_1)
    root_header_2 = build_header_hierarchy(list_of_headers_2)
    assert root_header_1 == root_header_2


def test_can_detect_mismatch() -> None:
    list_of_headers_1 = load_headers_from_file(
        Path("tests", "fixtures", "parser_test.docx")
    )
    list_of_headers_2 = load_headers_from_file(
        Path("tests", "fixtures", "parser_test_no_match.docx")
    )
    root_header_1 = build_header_hierarchy(list_of_headers_1)
    root_header_2 = build_header_hierarchy(list_of_headers_2)
    assert root_header_1 != root_header_2


def test_missing_headers() -> None:
    input_file = Path("tests", "fixtures", "parser_test.docx")
    list_of_headers = load_headers_from_file(input_file)
    root_header = build_header_hierarchy(list_of_headers)
    check_file = Path("tests", "fixtures", "parser_test_no_match.docx")
    list_of_headers_to_check = load_headers_from_file(check_file)
    check_header = build_header_hierarchy(list_of_headers_to_check)
    assert root_header.compare(check_header) == [
        "Heading 2 subheading 2 subsubheading 1"
    ]
