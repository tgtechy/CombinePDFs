import core.page_ops as p

print("MODULE FILE:", p.__file__)
print("HAS FUNCTION parse_page_range:", hasattr(p, "parse_page_range"))
print("DIR:", dir(p))