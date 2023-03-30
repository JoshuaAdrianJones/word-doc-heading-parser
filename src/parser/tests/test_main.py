from ..main import Header


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
