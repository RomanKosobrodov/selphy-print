from selphyprint.printing import get_printer_name, list_printers, get_printer_capabilities, CAPABILITIES

def test_printer_found():
    name = get_printer_name()
    assert name is not None

def test_list_printers():
    p = list_printers()
    assert len(p) > 0

def test_printer_capabilities():
    name = get_printer_name()
    c = get_printer_capabilities(name)
    for k in CAPABILITIES.keys():
        assert k in c.keys()


if __name__ == "__main__":
    test_printer_found()
    test_list_printers()


