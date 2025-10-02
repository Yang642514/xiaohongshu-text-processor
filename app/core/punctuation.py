import re

EN_TO_CN = {
    ",": "，",
    ".": "。",
    "?": "？",
    "!": "！",
    ":": "：",
    ";": "；",
    '"': '"',  # 将在引号处理逻辑中替换为中文双引号
    "'": "'",  # 将在引号处理逻辑中替换为中文单引号
    "(": "（",
    ")": "）",
}

def normalize_punctuation(text: str) -> str:
    # 先替换常见英文标点为中文标点
    for en, cn in EN_TO_CN.items():
        text = text.replace(en, cn)

    # 处理引号，将直引号替换为中文引号
    # 简化处理：成对出现的双引号替换为 “ ”，单引号替换为 ‘ ’
    def replace_quotes(s: str, quote: str, cn_open: str, cn_close: str) -> str:
        parts = s.split(quote)
        if len(parts) <= 1:
            return s
        out = []
        open_next = True
        for i, p in enumerate(parts):
            out.append(p)
            if i < len(parts) - 1:
                out.append(cn_open if open_next else cn_close)
                open_next = not open_next
        return "".join(out)

    text = replace_quotes(text, '"', '“', '”')
    text = replace_quotes(text, "'", '‘', '’')

    # 规范省略号：英文 ... -> 中文 ……
    text = re.sub(r"\.{3,}", "……", text)

    # 统一破折号: -- 或 —— -> ——
    text = re.sub(r"-{2,}", "——", text)

    return text