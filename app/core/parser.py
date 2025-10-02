import re
from typing import List, Tuple

# —— 尾段识别配置（默认值，允许被 configure_tail_filter 覆盖） ——
TAIL_FILTER_ENABLED = True
TAIL_PREFIXES = ["结尾", "总结", "写在最后", "最后", "文末", "结束语", "尾段", "尾声", "后记"]
TAIL_TITLE_KEYWORDS = [
    "总结", "结语", "写在最后", "最后", "结尾", "小结", "总结一下", "最后总结", "最后说两句",
    "文末", "写在文末", "尾部", "尾段", "结束语", "END", "The End", "P.S.", "PS",
    "尾声", "后记", "作者的话", "免责声明", "广告", "活动", "拓展阅读", "扩展阅读", "参考资料"
]
CTA_KEYWORDS = [
    "关注", "关注我", "点赞", "点个赞", "收藏", "收藏一下", "评论", "留言", "转发", "在看",
    "一键三连", "三连", "投币", "私信", "下单", "购买", "店铺", "链接", "主页", "主页链接",
    "点我", "扫码", "优惠", "折扣", "团购", "预约", "报名", "加微信", "VX", "公众号",
    "粉丝群", "欢迎关注", "感谢观看", "谢谢", "求关注", "求点赞", "求收藏", "评论区", "看评论",
    "看更多", "更多内容", "更多干货", "更多技巧", "更多案例", "下方链接"
]
TAIL_SHORT_THRESHOLD = 120

def configure_tail_filter(settings: dict):
    """根据设置更新尾段过滤相关配置。"""
    global TAIL_FILTER_ENABLED, TAIL_PREFIXES, TAIL_TITLE_KEYWORDS, CTA_KEYWORDS, TAIL_SHORT_THRESHOLD
    TAIL_FILTER_ENABLED = bool(settings.get("tail_filter_enabled", True))
    prefixes = settings.get("tail_prefixes")
    if isinstance(prefixes, list) and prefixes:
        TAIL_PREFIXES = prefixes
    keywords = settings.get("tail_keywords")
    if isinstance(keywords, list) and keywords:
        TAIL_TITLE_KEYWORDS = keywords
    ctas = settings.get("cta_keywords")
    if isinstance(ctas, list) and ctas:
        CTA_KEYWORDS = ctas
    try:
        TAIL_SHORT_THRESHOLD = int(settings.get("tail_short_threshold", 120))
    except Exception:
        TAIL_SHORT_THRESHOLD = 120

def is_tail_section(title: str, content: str) -> bool:
    """判断一个分论点是否更像是文案收尾/行动召唤段落。
    规则（启发式）：
    - 标题包含尾段关键词（如“总结/写在最后/最后/END/后记/免责声明”等），且内容很短或包含CTA关键词。
    - 或内容中包含较多CTA关键词（计数≥2）且整体较短（<120字符）。
    - 仅用于过滤末尾分论点，避免误判。
    """
    t = (title or "").strip()
    c = (content or "").strip()
    t_lower = t.lower()
    c_lower = c.lower()

    title_hit = any(k.lower() in t_lower for k in TAIL_TITLE_KEYWORDS)
    cta_count = sum(1 for k in CTA_KEYWORDS if k.lower() in c_lower)
    short = len(c) < TAIL_SHORT_THRESHOLD

    if title_hit and (cta_count >= 1 or short):
        return True
    if cta_count >= 2 and short:
        return True
    # 若标题是单个符号或表情开头的短句，且内容几乎为空，也视为尾段
    if re.match(r"^[\-•*#\u2014]+", t) and len(c) < 30:
        return True
    return False


def extract_title_and_points(raw: str) -> Tuple[str, List[Tuple[str, str]]]:
    """
    输入整体文案文本，返回标题与分论点列表。
    分论点列表元素为 (point_title, point_content)，保持输入顺序。
    简化策略：
    - 第一行作为标题
    - 其余行中，匹配可能的分论点标题行：
      * 中文序号：一、二、三、...；或者（1）（2）
      * 数字序号：1. 2. 3. 或者 1、 2、
      * 列点符号：- • *
    - 标题行后续非标题行拼接为该分论点内容，直到下一个标题行。
    """
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", []

    title = lines[0]
    content_lines = lines[1:]

    point_patterns = [
        r"^论点[一二三四五六七八九十百]+[：:、.，]?",  # 论点一 论点二
        r"^(?:第?([一二三四五六七八九十百]+)章?|([一二三四五六七八九十百]+))、",  # 一、 二、
        r"^(第\d+点)",  # 第1点 第2点
        r"^(\d+)[\.、]",  # 1. 1、
        r"^[\-•*]",  # - • *
        r"^（?\d+）",  # （1）(1)
    ]
    point_regex = re.compile("|".join(point_patterns))

    points: List[Tuple[str, str]] = []
    current_point_title = None
    current_point_content: List[str] = []

    def flush_current():
        nonlocal current_point_title, current_point_content
        if current_point_title is not None:
            points.append((current_point_title, "\n".join(current_point_content).strip()))
        current_point_title = None
        current_point_content = []

    tail_started = False

    for line in content_lines:
        # 一旦检测到尾段前缀，标记后续全部作为尾段，后续统一剥离（受开关控制）
        if TAIL_FILTER_ENABLED and not tail_started:
            l_strip = line.strip()
            if any(l_strip.startswith(p) for p in TAIL_PREFIXES):
                tail_started = True

        if TAIL_FILTER_ENABLED and tail_started:
            # 收集到 current_point_content，稍后将整体视为尾段并剥离
            if current_point_title is None:
                current_point_title = (line.strip() or "尾段")
            current_point_content.append(line)
            continue

        if point_regex.search(line):
            # 新的分论点开始
            flush_current()
            current_point_title = line
        else:
            # 内容行
            if current_point_title is None:
                # 如果没有明确的分论点标题，则将第一段作为分论点1
                current_point_title = "分论点1"
            current_point_content.append(line)

    flush_current()

    # 若检测到尾段前缀，则整体剥离末尾块（受开关控制）
    if TAIL_FILTER_ENABLED and tail_started and points:
        last_t, last_c = points[-1]
        # 如果最后一个块包含我们收集到的尾段内容，则移除
        if is_tail_section(last_t, last_c) or any(last_c.startswith(p) for p in TAIL_PREFIXES):
            points.pop()

    # 额外过滤：连续多个尾段（CTA 尾巴等）（受开关控制）
    if TAIL_FILTER_ENABLED:
        while points:
            last_t, last_c = points[-1]
            if is_tail_section(last_t, last_c):
                points.pop()
            else:
                break

    return title, points

def format_paragraphs(text: str, max_lines: int = 0, max_chars: int = 45) -> str:
    """按句号换行，并限制段落最多 max_lines 行。
    - 句子边界：中文句号 "。" 与英文句点 "."，保留标点。
    - 当 max_chars <= 0 时，不进行按字符切分，每句占一行。
    - max_lines<=0 表示不限制行数；>0 时仅保留前 max_lines 行。
    - 不删除内容，完整保留文本。
    """
    clean = re.sub(r"\s+", " ", text.strip())
    if not clean:
        return ""

    # 按句号分句（保留标点）
    parts = []
    buf = []
    for ch in clean:
        buf.append(ch)
        if ch in "。.":
            parts.append("".join(buf))
            buf = []
    if buf:
        parts.append("".join(buf))

    lines: List[str] = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if max_chars and max_chars > 0:
            i = 0
            while i < len(p):
                lines.append(p[i:i+max_chars])
                i += max_chars
        else:
            # 不做字符折行：整句作为一行
            lines.append(p)
    if max_lines and max_lines > 0:
        lines = lines[:max_lines]
    return "\n".join(lines)


def render_processed_template(points: List[Tuple[str, str]]) -> str:
    numerals = ["一","二","三","四","五","六","七","八","九","十","十一","十二","十三","十四","十五","十六","十七","十八","十九","二十"]
    parts = []
    for idx, (title, content) in enumerate(points):
        num_cn = numerals[idx] if idx < len(numerals) else f"{idx+1}"
        parts.append(f"分论点{num_cn}：\n{title}\n分论点{num_cn}内容：\n{content}")
    return "\n\n".join(parts)


def parse_processed_template(text: str) -> List[Tuple[str, str]]:
    """
    解析渲染后的处理模板文本，格式类似：
    分论点一：\n<标题>\n分论点一内容：\n<多行内容...>\n\n分论点二：...

    解析规则：
    - 识别以“分论点…：”结尾的提示行作为块开始；下一行是标题。
    - 识别“分论点…内容：”作为内容起始；直到下一个“分论点…：”或文本结束为止均为内容。
    - 保留内容中的所有换行，不丢弃任何文本。
    """
    lines = [l.rstrip("\n") for l in text.splitlines()]
    pairs: List[Tuple[str, str]] = []
    i = 0
    start_re = re.compile(r"^分论点[一二三四五六七八九十\d]+：$")
    content_re = re.compile(r"^分论点[一二三四五六七八九十\d]+内容：$")

    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if start_re.match(line):
            # 读取标题行
            i += 1
            # 跳过空行
            while i < len(lines) and not lines[i].strip():
                i += 1
            if i >= len(lines):
                break
            title = lines[i].strip()

            # 查找内容提示行
            i += 1
            while i < len(lines) and not content_re.match(lines[i].strip()):
                # 容错：如果中间夹杂空行或其他文本，继续向后
                i += 1
            # 跳过“内容：”提示行
            if i < len(lines) and content_re.match(lines[i].strip()):
                i += 1

            # 收集内容直到下一个分论点块开始或结束
            content_buf: List[str] = []
            while i < len(lines) and not start_re.match(lines[i].strip()):
                content_buf.append(lines[i])
                i += 1
            content = "\n".join(content_buf).strip()
            pairs.append((title, content))
        else:
            i += 1
    return pairs