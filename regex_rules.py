REGEX_ORDER = [
    ("fixed", r'^第[\d一二三四五六七八九十百]+章[ 　\t]*'),
    ("fixed", r'^第[\d一二三四五六七八九十百]+节[ 　\t]*'),
    ("fixed", r'^第[\d一二三四五六七八九十百]+条[ 　\t]*'),
    ("numeric_dotted", None),
    ("fixed", r'^[（(]\s*\d+\s*[)）]'),
    ("fixed", r'^[（(]\s*[一二三四五六七八九十百]+\s*[)）]'),
    ("fixed", r'^\d+\s*[)）]'),
    ("fixed", r'^[一二三四五六七八九十百]+、\s*'),
]

NEGATIVE_PATTERNS = [
    r'^\s*\d+\.\s*(cop[0-4]|sp|sol|un)',   # 表注里的编号行（2. cop4…)
]

