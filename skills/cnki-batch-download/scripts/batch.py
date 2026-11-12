#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""cnki-batch-download helper.

Manifest parsing, dedupe state, download verification and reporting for batch
CNKI downloads driven by the chrome-devtools MCP browser.

Subcommands
-----------
parse     manifest file -> items.json
list      print items with current status
next      print next pending item (JSON)
mark      update one item status
scan      list files in a directory newer than a marker (download detection)
existing  match pending items against files already in a directory
report    write report.md + failed.md and print a summary table

State lives in <workdir>/state.json; the parsed manifest in <workdir>/items.json.
Settings live in <skill_dir>/config.json (created by `init`).
Python: stdlib only, except .xlsx input which needs openpyxl.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from difflib import SequenceMatcher

CJK = "\u4e00-\u9fff"
NORM_RE = re.compile(r"[^\w" + CJK + r"]+", re.UNICODE)
YEAR_RE = re.compile(r"(?:19|20)\d{2}")
URL_RE = re.compile(r"^https?://\S+", re.IGNORECASE)
CNKI_URL_RE = re.compile(r"https?://[^\s]*cnki\.net[^\s]*", re.IGNORECASE)

# GB/T 7714 style:  [1] 作者, 作者. 题名[J]. 刊名, 2020, 43(5): 1-10.
GBT_RE = re.compile(
    r"^\s*(?:\[\d+\]|\d+[.、])?\s*(?P<authors>[^.]{0,200}?)[.．]\s*"
    r"(?P<title>.{2,300}?)\s*\[(?P<kind>[JDCMSRNPZ])\]"
    r"(?P<rest>.{0,300})$"
)
TITLE_FIRST_RE = re.compile(
    r"^\s*(?:\[\d+\]|\d+[.、])?\s*(?P<title>.{2,300}?)\s*"
    r"\[(?P<kind>[JDCMSRNPZ])\]\s*[.．]?\s*(?P<rest>.{0,300})$"
)
# APA style:  Author, A. (2020). Title of the work. Journal, 43(5), 1-10.
APA_RE = re.compile(
    r"^\s*(?P<authors>.{0,200}?)\((?P<year>(?:19|20)\d{2})[a-z]?\)[.．]\s*"
    r"(?P<title>.{2,300}?)[.．]\s(?P<rest>.{0,300})$"
)

STATUSES = ("pending", "ok", "failed", "skipped")
HEADER_HINTS = {
    "title": ("题名", "标题", "篇名", "论文", "title", "文献名"),
    "authors": ("作者", "author", "作 者"),
    "year": ("年份", "年", "year", "发表"),
    "url": ("链接", "网址", "url", "地址", "详情"),
    "doi": ("doi",),
}


# ---------------------------------------------------------------- utilities


def norm(text: str) -> str:
    """Fold to lowercase word characters only; used for keys and matching."""
    return NORM_RE.sub("", (text or "").lower())


def key_for(title: str, url: str = "") -> str:
    base = norm(title) or norm(url)
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:12]


def _title_variants(title: str) -> list:
    """知网下载的文件名常丢掉副标题（——之后那段），或加 _作者，这里都算上。"""
    t = (title or "").strip()
    out = [t]
    for sep in ("——", "—", "--", "：", ":", "（", "("):
        if sep in t:
            out.append(t.split(sep)[0].strip())
    return [x for x in out if len(x) >= 4]


def _stem_variants(stem: str) -> list:
    out = [stem]
    for sep in ("_", "（", "(", "-"):
        if sep in stem:
            out.append(stem.split(sep)[0].strip())
    return [x for x in out if len(x) >= 4]


def similar(a: str, b: str) -> float:
    na, nb = norm(a), norm(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    if len(na) >= 8 and (na in nb or nb in na):
        return 0.99
    return SequenceMatcher(None, na, nb).ratio()


def guess_year(text: str) -> str:
    m = YEAR_RE.search(text or "")
    return m.group(0) if m else ""


def cleanup_title(text: str) -> str:
    t = (text or "").strip()
    t = re.sub(r"^\s*\[\d+\]\s*", "", t)
    t = re.sub(r"\s*\[(?:J|D|C|M|N|R|P|Z|EB/OL)\]\s*$", "", t)
    t = t.strip(" .。;；,，:：\"'“”‘’")
    return t


def workdir_paths(workdir: str) -> tuple[str, str]:
    return (
        os.path.join(workdir, "items.json"),
        os.path.join(workdir, "state.json"),
    )


def load_json(path: str, default):
    if not os.path.exists(path):
        return default
    with io.open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def save_json(path: str, data) -> None:
    """Atomic-ish write that survives OneDrive/AV briefly locking the target."""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    last = None
    for attempt in range(6):
        try:
            os.replace(tmp, path)
            return
        except OSError as exc:  # WinError 5/32 on OneDrive or AV scanning
            last = exc
            time.sleep(0.15 * (attempt + 1))
    # last resort: write in place, accepting a non-atomic window
    try:
        with io.open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        try:
            os.remove(tmp)
        except OSError:
            pass
        return
    except OSError:
        raise last if last else OSError("could not write " + path)


# ---------------------------------------------------------------- parsing


def row_to_item(title: str, authors: str = "", year: str = "",
                url: str = "", doi: str = "", kind: str = "text") -> dict | None:
    title = cleanup_title(title)
    url = (url or "").strip()
    if not title and not url:
        return None
    if not year:
        year = guess_year(title) or guess_year(authors)
    return {
        "key": key_for(title, url),
        "title": title,
        "authors": (authors or "").strip(),
        "year": (year or "").strip(),
        "url": url,
        "doi": (doi or "").strip(),
        "kind": kind,
        "status": "pending",
        "file": "",
        "reason": "",
        "updated": "",
    }


def _split_top(text: str, sep: str) -> list:
    """只在书名号/括号外面切分——《黑神话：悟空》这种题名不能被切成两截。"""
    out, buf, depth = [], [], 0
    for ch in text:
        if ch in "《（(【[":
            depth += 1
        elif ch in "》）)】]":
            depth = max(0, depth - 1)
        if ch == sep and depth == 0:
            out.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    out.append("".join(buf))
    return out


def parse_free_line(line: str) -> dict | None:
    """One reference line in any of the supported plain-text shapes."""
    raw = line.strip().strip("\ufeff")
    if not raw or raw.startswith("#"):
        return None

    # bare URL / CNKI link
    urls = CNKI_URL_RE.findall(raw)
    if urls and len(norm(raw.replace(urls[0], ""))) <= 2:
        return row_to_item(title="", url=urls[0], kind="url")

    # markdown table row: | title | authors | year |
    if raw.startswith("|"):
        cells = [c.strip() for c in raw.strip("|").split("|")]
        if all(re.fullmatch(r":?-{2,}:?", c or "-") for c in cells):
            return None
        cells = [c for c in cells if c]
        if not cells:
            return None
        if any(h in cells[0].lower() for h in HEADER_HINTS["title"]):
            return None
        url = next((c for c in cells if URL_RE.match(c)), "")
        title = cells[0]
        authors = cells[1] if len(cells) > 1 and not URL_RE.match(cells[1]) else ""
        year = next((c for c in cells if re.fullmatch(r"(?:19|20)\d{2}", c)), "")
        return row_to_item(title, authors, year, url, kind="table")

    m = GBT_RE.match(raw)
    if m:
        year = guess_year(m.group("rest"))
        return row_to_item(m.group("title"), m.group("authors"), year,
                           urls[0] if urls else "", kind="gbt")

    # authorless GB/T:  题名[J]. 刊名, 2020, 40(2): 1-10.
    m = TITLE_FIRST_RE.match(raw)
    if m:
        return row_to_item(m.group("title"), "", guess_year(m.group("rest")),
                           urls[0] if urls else "", kind="titlefirst")

    m = APA_RE.match(raw)
    if m:
        return row_to_item(m.group("title"), m.group("authors"),
                           m.group("year"),
                           urls[0] if urls else "", kind="apa")

    # explicit separators: title | authors | year, title \t authors, title::url
    for sep in ("\t", "|", "：", "::"):
        if sep in raw:
            parts = [p.strip() for p in _split_top(raw, sep) if p.strip()]
            if len(parts) >= 2:
                url = next((p for p in parts if URL_RE.match(p)), "")
                title = parts[0]
                rest = [p for p in parts[1:] if not URL_RE.match(p)]
                authors = rest[0] if rest else ""
                year = next((p for p in rest if re.fullmatch(r"(?:19|20)\d{2}", p)), "")
                return row_to_item(title, authors, year, url, kind="sep")

    # fallback: "作者. 题名. 刊名, 2020" or a bare title
    year = guess_year(raw)
    if ". " in raw or "．" in raw:
        segs = [s.strip() for s in re.split(r"[.．]\s+", raw) if s.strip()]
        if len(segs) >= 2:
            # prefer the first segment after the author block that carries no year
            idx = next(
                (i for i, seg in enumerate(segs[1:], start=1)
                 if not YEAR_RE.search(seg) and not re.search(r"[,，]\s*\d", seg)),
                1,
            )
            return row_to_item(segs[idx], segs[0], year, urls[0] if urls else "",
                               kind="ref")
    return row_to_item(raw, "", year, urls[0] if urls else "", kind="line")


def detect_columns(header: list[str]) -> dict:
    mapping = {}
    lowered = [(i, (c or "").strip().lower()) for i, c in enumerate(header)]
    for field, hints in HEADER_HINTS.items():
        for i, cell in lowered:
            if any(h.lower() in cell for h in hints):
                mapping[field] = i
                break
    return mapping


def parse_table(rows: list[list[str]]) -> list[dict]:
    rows = [[("" if c is None else str(c)).strip() for c in r] for r in rows]
    rows = [r for r in rows if any(r)]
    if not rows:
        return []
    mapping = detect_columns(rows[0])
    start = 1 if mapping else 0
    if not mapping:
        mapping = {"title": 0, "authors": 1, "year": 2}

    items = []
    for r in rows[start:]:
        def get(field):
            i = mapping.get(field)
            return r[i] if i is not None and i < len(r) else ""
        item = row_to_item(get("title"), get("authors"), get("year"),
                           get("url"), get("doi"), kind="table")
        if item:
            items.append(item)
    return items


def parse_xlsx(path: str) -> list[dict]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover
        raise SystemExit(
            "openpyxl is missing from this interpreter.\n"
            "Use the bundled runtime: <DSH_HOME>/dsh-runtimes/dsh-primary-runtime/"
            "dependencies/python/python.exe\n"
            "Original error: %s" % exc
        )
    wb = load_workbook(path, read_only=True, data_only=True)
    items: list[dict] = []
    for sheet in wb.worksheets:
        rows = [list(r) for r in sheet.iter_rows(values_only=True)]
        items.extend(parse_table(rows))
    wb.close()
    return items


def parse_csv(path: str) -> list[dict]:
    with io.open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
        sample = fh.read(8192)
        fh.seek(0)
        delim = ","
        try:
            delim = csv.Sniffer().sniff(sample, delimiters=",\t;|").delimiter
        except csv.Error:
            if "\t" in sample:
                delim = "\t"
        rows = list(csv.reader(fh, delimiter=delim))
    return parse_table(rows)


def parse_text(path: str) -> list[dict]:
    with io.open(path, "r", encoding="utf-8-sig", errors="replace") as fh:
        lines = fh.read().splitlines()
    items = [i for i in (parse_free_line(ln) for ln in lines) if i]
    return items


def parse_json(path: str) -> list[dict]:
    data = load_json(path, [])
    if isinstance(data, dict):
        data = data.get("items", [])
    items = []
    for row in data:
        item = row_to_item(row.get("title", ""), row.get("authors", ""),
                           row.get("year", ""), row.get("url", ""),
                           row.get("doi", ""), kind="json")
        if item:
            items.append(item)
    return items


def parse_manifest(path: str) -> list[dict]:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        items = parse_xlsx(path)
    elif ext == ".csv":
        items = parse_csv(path)
    elif ext == ".json":
        items = parse_json(path)
    else:
        items = parse_text(path)

    seen: dict[str, int] = {}
    unique = []
    for item in items:
        k = item["key"]
        if k in seen:
            continue
        seen[k] = 1
        unique.append(item)
    for idx, item in enumerate(unique, start=1):
        item["n"] = idx
    return unique


# ---------------------------------------------------------------- state


def load_state(workdir: str) -> tuple[dict, list[dict]]:
    """items.json is authoritative; state.json is a derived mirror.

    Keeping the mirror read-only avoids an older state entry clobbering a newer
    status written by `mark`.
    """
    items_path, state_path = workdir_paths(workdir)
    state = load_json(state_path, {"items": {}, "created": "", "download_dir": ""})
    state.setdefault("items", {})
    state.setdefault("created", "")
    state.setdefault("download_dir", "")
    items = load_json(items_path, [])
    return state, items


def sync_state(workdir: str) -> tuple[dict, list[dict]]:
    items_path, state_path = workdir_paths(workdir)
    state, items = load_state(workdir)
    state["items"] = {}
    for item in items:
        state["items"][item["key"]] = {
            "n": item.get("n"),
            "title": item.get("title", ""),
            "authors": item.get("authors", ""),
            "year": item.get("year", ""),
            "url": item.get("url", ""),
            "status": item.get("status", "pending"),
            "file": item.get("file", ""),
            "reason": item.get("reason", ""),
            "updated": item.get("updated", ""),
        }
    if not state.get("created"):
        state["created"] = time.strftime("%Y-%m-%d %H:%M:%S")
    save_json(state_path, state)
    return state, items


def find_item(items: list[dict], key: str = "", n: int = 0,
              title: str = "") -> dict | None:
    if key:
        for item in items:
            if item["key"] == key:
                return item
    if n:
        for item in items:
            if item.get("n") == n:
                return item
    if title:
        best, score = None, 0.0
        for item in items:
            s = similar(title, item["title"])
            if s > score:
                best, score = item, s
        if best and score >= 0.8:
            return best
    return None


# ---------------------------------------------------------------- commands


# --------------------------------------------------------------- settings
#
# 目标：任何学校的用户都能用。用户只需要提供「一条能打开知网的网址」，
# 脚本自己判断走哪种机构访问方式，生成 config.json。

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 配置文件放用户主目录：skill 安装目录常常是只读的，不能往里写
USER_CONFIG = os.path.join(os.path.expanduser("~"), ".cnki-batch-download", "config.json")
LEGACY_CONFIG = os.path.join(SKILL_DIR, "config.json")
CWD_CONFIG = os.path.join(os.getcwd(), ".cnki-batch-download", "config.json")
DEFAULT_CONFIG = USER_CONFIG

# access.mode 取值与含义（写给用户看的大白话）
MODE_HELP = {
    "host": "学校 WebVPN 把域名编进网址，形如 https://kns-cnki-net-443.v.学校域名/",
    "prefix": "学校 VPN 在网址前面加一段加密串，形如 https://vpn.学校域名/https/<串>/kns8s/",
    "ezproxy": "图书馆代理，形如 https://libproxy.学校域名/login?url=知网网址",
    "direct": "直接访问知网，不改网址（个人账号 / 校园网 IP / 已登录 CARSI 都走这条）",
}

LOGIN_HELP = {
    "personal": {
        "名称": "个人账号（自己注册，按篇付费）",
        "说明": "没有单位，用手机号/邮箱注册的知网账号，下载时按篇扣余额",
        "要网址": "",
        "下一步": "在浏览器打开 {entry} ，用手机号/邮箱登录（没有账号先注册）。"
                  "个人账号下载按篇计费，遇到收费提示是正常的，是否付费由你决定。"
                  "登录后回一句「继续」。",
    },
    "institution": {
        "名称": "机构用户登录（知网官方入口）",
        "说明": "单位/学校给了知网机构账号，或在学校 IP 内，用页面右上「机构登录」",
        "要网址": "",
        "下一步": "在浏览器打开 {entry} ，点右上角「机构登录」，用单位账号或校园网 IP 登录。"
                  "页头出现单位名即成功。登录后回一句「继续」。",
    },
    "webvpn": {
        "名称": "学校 WebVPN（域名编进网址）",
        "说明": "要先进学校 VPN 门户，知网网址会变成 kns-cnki-net-443.v.学校域名 这种样子",
        "要网址": "在 VPN 里打开任意一条知网页面，把地址栏那条 xxx-cnki-net-443.v.学校域名 的网址粘过来。",
        "下一步": "先在浏览器登录学校 VPN 门户，确认 {entry} 能打开且页头显示学校名，"
                  "然后回一句「继续」。",
    },
    "prefix_vpn": {
        "名称": "学校路径前缀 VPN",
        "说明": "网址前面多一段加密路径，形如 vpn.学校域名/https/<一串>/kns8s/",
        "要网址": "在 VPN 里打开任意一条知网页面，把地址栏那条 /https/<一串>/ 形式的网址粘过来。",
        "下一步": "先在浏览器登录学校 VPN 门户，确认 {entry} 能打开且页头显示学校名，"
                  "然后回一句「继续」。",
    },
    "ezproxy": {
        "名称": "图书馆 ezproxy 代理",
        "说明": "从图书馆网站点进知网，网址形如 libproxy.学校域名/login?url=...",
        "要网址": "从图书馆网站点进知网，把地址栏那条 libproxy.学校域名/login?url= 的网址粘过来。",
        "下一步": "先在浏览器登录图书馆账号，确认 {entry} 能打开且页头显示单位名，"
                  "然后回一句「继续」。",
    },
}

CONFIG_TEMPLATE = {
    "version": 2,
    "school": {"name": "", "domain": ""},
    "access": {
        "mode": "direct",
        "school_domain": "",
        "prefix": "",
        "ezproxy": "",
        "entry_url": "",
    },
    "download_dir": "",
    "thresholds": {"match": 0.9, "existing": 0.82, "page_title": 0.75},
}

CNKI_HOST_RE = re.compile(
    r"^https?://([a-z0-9.-]*cnki[a-z0-9.-]*)\.net(?::(\d+))?(/.*)?$", re.I)
# 学校 WebVPN 主机名：kns-cnki-net-443.v.xx.edu.cn
WEBVPN_HOST_RE = re.compile(
    r"^https?://(?P<label>[a-z0-9-]+)\.v\.(?P<school>[a-z0-9.-]+)(?::\d+)?(/.*)?$",
    re.I)
# 站点内部路径起点，用来切出 VPN 前缀
INNER_MARKERS = ("/kns8s/", "/kcms2/", "/kns/", "/dm8/", "/bar/", "/grid2008/")


def config_path(args) -> str:
    """--config > 环境变量 > 老位置(skill 目录) > 主目录 > 当前目录，取第一个存在的。"""
    explicit = getattr(args, "config", "") or ""
    if explicit:
        return explicit
    env = os.environ.get("CNKI_BATCH_CONFIG", "")
    if env:
        return env
    for cand in (LEGACY_CONFIG, USER_CONFIG, CWD_CONFIG):
        if os.path.exists(cand):
            return cand
    return USER_CONFIG


def load_config(path: str) -> dict:
    cfg = json.loads(json.dumps(CONFIG_TEMPLATE))
    if path and os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                disk = json.load(f)
        except (OSError, ValueError):
            disk = {}
        for k, v in disk.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    return cfg


def save_config(path: str, cfg: dict) -> None:
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    save_json(path, cfg)


def infer_access(url: str) -> dict:
    """从一条能打开的知网网址反推访问方式。返回 {mode, ...}，无法判断时 mode=''。"""
    url = (url or "").strip()
    if not url:
        return {"mode": ""}

    # 图书馆代理：.../login?url=<编码后的知网>
    m = re.match(r"^(https?://[^/]+/[^?#]*?[?&]url=)(.+)$", url, re.I)
    if m and re.search(r"libproxy|ezproxy|proxy|lib", url, re.I):
        return {"mode": "ezproxy", "ezproxy": m.group(1)}

    # 学校 WebVPN 主机名映射
    m = WEBVPN_HOST_RE.match(url)
    if m and "cnki" in m.group("label"):
        return {"mode": "host", "school_domain": m.group("school").lower()}

    # VPN 路径前缀：/https/<token>/kns8s/...
    for marker in INNER_MARKERS:
        i = url.find(marker)
        if i > 0 and not CNKI_HOST_RE.match(url):
            return {"mode": "prefix", "prefix": url[:i]}

    # 原生知网地址 = 直连
    if CNKI_HOST_RE.match(url):
        return {"mode": "direct"}

    return {"mode": ""}


def rewrite_url(url: str, cfg: dict) -> tuple[str, bool]:
    """把知网原生网址换成本校可访问的网址。-> (url, 是否支持)"""
    if not url:
        return url, True
    access = cfg.get("access") or {}
    mode = access.get("mode") or "direct"
    if mode == "direct":
        return url, True
    m = CNKI_HOST_RE.match(url)
    if not m:
        return url, True          # 非知网网址原样保留
    inner = m.group(3) or "/"

    if mode == "host":
        school = (access.get("school_domain") or "").strip()
        sub = m.group(1).replace(".cnki", "").replace(".", "-").strip("-") or "www"
        port = m.group(2) or "443"
        if not school:
            return url, False
        # 已经是本校代理域就不再改
        if url.startswith("https://") and f".v.{school}" in url:
            return url, True
        return f"https://{sub}-cnki-net-{port}.v.{school}{inner}", True

    if mode == "prefix":
        prefix = (access.get("prefix") or "").rstrip("/")
        if not prefix:
            return url, False
        return prefix + inner, True

    if mode == "ezproxy":
        base = access.get("ezproxy") or ""
        if not base:
            return url, False
        if url.startswith(base):
            return url, True
        return base + urllib.parse.quote(url, safe=""), True

    return url, False


def _ask(prompt: str, default: str = "") -> str:
    suffix = f" [回车={default}]" if default else ""
    try:
        got = input(f"{prompt}{suffix}: ").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return default
    return got or default


LOGIN_TO_MODE = {
    "personal": "direct",
    "institution": "direct",
    "webvpn": "host",
    "prefix_vpn": "prefix",
    "ezproxy": "ezproxy",
}


def cmd_init(args) -> int:
    """一次性配置。第一步永远是先问「你怎么登录知网」：
       personal    个人账号（自己充值，可能按篇付费）
       institution 知网官方机构用户登录（页面右上「机构登录」）
       webvpn      学校 WebVPN，网址形如 kns-cnki-net-443.v.学校域名
       prefix_vpn  学校路径前缀 VPN，网址形如 vpn.学校域名/https/<串>/kns8s
       ezproxy     图书馆 ezproxy，网址形如 libproxy.学校域名/login?url=...
    """
    path = config_path(args)
    cfg = load_config(path)
    access = cfg["access"]

    login = (getattr(args, "login", "") or "").strip()
    from_url = getattr(args, "from_url", "") or ""

    # 1) 先定登录方式
    if not login and args.mode:
        login = next((k for k, v in LOGIN_TO_MODE.items() if v == args.mode), "")
    if not login and from_url:
        guess_mode = infer_access(from_url).get("mode")
        login = next((k for k, v in LOGIN_TO_MODE.items() if v == guess_mode), "")
    if not login and sys.stdin.isatty():
        print("先确认一件事：你在浏览器里是怎么登录知网的？")
        for i, (k, v) in enumerate(LOGIN_HELP.items(), 1):
            print(f"  {i}. {v['名称']} —— {v['说明']}")
        print("  不确定就选 1 或 2；之后随时可以用 init --login 改。")
        pick = _ask("请输入数字", "1")
        keys = list(LOGIN_HELP)
        if pick.isdigit() and 1 <= int(pick) <= len(keys):
            login = keys[int(pick) - 1]
        elif pick in LOGIN_HELP:
            login = pick
    if not login:
        print(json.dumps({
            "error": "还没确定登录方式，配置未保存",
            "怎么填": "init --login personal | institution | webvpn | prefix_vpn | ezproxy",
            "登录方式": {k: v["名称"] for k, v in LOGIN_HELP.items()},
        }, ensure_ascii=False, indent=2))
        return 2
    if login not in LOGIN_TO_MODE:
        print(json.dumps({"error": "--login 只能是: " + ", ".join(LOGIN_TO_MODE),
                          "登录方式": {k: v["名称"] for k, v in LOGIN_HELP.items()}},
                         ensure_ascii=False, indent=2))
        return 2

    want_mode = LOGIN_TO_MODE[login]

    # 2) 只有机构类通道才需要问学校网址
    if want_mode != "direct":
        if not from_url and sys.stdin.isatty():
            print()
            print(f"你选的是「{LOGIN_HELP[login]['名称']}」。")
            print(LOGIN_HELP[login]["要网址"])
            from_url = _ask("知网网址")
        if not from_url and not (args.school_domain or args.prefix or args.ezproxy):
            print(json.dumps({
                "error": "这种方式必须先拿到你们学校的知网网址，配置未保存",
                "怎么填": 'init --login ' + login + ' --from-url "<浏览器地址栏里那条网址>"',
                "为什么": LOGIN_HELP[login]["要网址"],
            }, ensure_ascii=False, indent=2))
            return 2
        if from_url:
            guess = infer_access(from_url)
            if not guess.get("mode"):
                print(json.dumps({
                    "error": "认不出这个网址属于哪种访问方式",
                    "got": from_url,
                    "hint": "改用手动参数：--mode host --school-domain 学校域名，"
                            "或 --mode prefix --prefix VPN前缀，或 --mode direct。",
                    "modes": MODE_HELP,
                }, ensure_ascii=False, indent=2))
                return 2
            if guess["mode"] != want_mode:
                print(json.dumps({
                    "提醒": "网址识别为 " + guess["mode"] + "，和你选的 " + want_mode
                            + " 不一致，按网址的识别结果设置",
                }, ensure_ascii=False, indent=2))
            access.update({k: v for k, v in guess.items() if k != "mode"})
            access["mode"] = guess["mode"]
            access["entry_url"] = from_url
            print(json.dumps({"识别结果": MODE_HELP[guess["mode"]], "mode": guess["mode"]},
                             ensure_ascii=False, indent=2))
        else:
            access["mode"] = want_mode
    else:
        access["mode"] = "direct"
        access["entry_url"] = ("https://www.cnki.net/" if login == "personal"
                               else "https://kns.cnki.net/")

    access["login_type"] = login

    # 3) 显式参数覆盖
    if args.mode:
        if args.mode not in MODE_HELP:
            print(json.dumps({"error": "mode 只能是: " + ", ".join(MODE_HELP)},
                             ensure_ascii=False))
            return 2
        access["mode"] = args.mode
    if args.school_name:
        cfg["school"]["name"] = args.school_name
    if args.school_domain:
        cfg["school"]["domain"] = args.school_domain
        access["school_domain"] = args.school_domain
    if args.prefix:
        access["prefix"] = args.prefix.rstrip("/")
    if args.ezproxy:
        access["ezproxy"] = args.ezproxy
    if args.entry_url:
        access["entry_url"] = args.entry_url
    if args.download_dir:
        cfg["download_dir"] = args.download_dir

    # 4) 补默认值 + 校验
    if not access.get("entry_url"):
        d = access.get("school_domain")
        if access["mode"] == "host" and d:
            access["entry_url"] = f"https://www-cnki-net-443.v.{d}/"
        elif access["mode"] == "direct":
            access["entry_url"] = "https://www.cnki.net/"
    if not cfg.get("download_dir"):
        cfg["download_dir"] = os.path.join(os.path.expanduser("~"), "Downloads")

    problems = []
    if access["mode"] == "host" and not access.get("school_domain"):
        problems.append("host 模式缺少学校域名（例：xx.edu.cn），加 --school-domain")
    if access["mode"] == "prefix" and not access.get("prefix"):
        problems.append("prefix 模式缺少 VPN 前缀，加 --prefix")
    if access["mode"] == "ezproxy" and not access.get("ezproxy"):
        problems.append("ezproxy 模式缺少代理前缀，加 --ezproxy")
    if problems:
        print(json.dumps({"error": "配置不完整", "problems": problems},
                         ensure_ascii=False, indent=2))
        return 2

    note = ""
    try:
        save_config(path, cfg)
    except OSError:
        # 主目录或 skill 目录不可写（沙箱/只读安装）时退到当前目录
        path = CWD_CONFIG
        save_config(path, cfg)
        note = "主目录不可写，配置存在当前目录；下次在别处运行请加 --config"
    print(json.dumps({
        "saved": path,
        "提示": note,
        "登录方式": LOGIN_HELP[login]["名称"],
        "mode": access["mode"],
        "说明": (MODE_HELP[access["mode"]] if access["mode"] != "direct"
                 else "直接访问知网，不改网址"),
        "entry_url": access.get("entry_url"),
        "download_dir": cfg["download_dir"],
        "下一步": LOGIN_HELP[login]["下一步"].format(entry=access.get("entry_url")),
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_config(args) -> int:
    path = config_path(args)
    cfg = load_config(path)
    print(json.dumps({"path": path, "exists": os.path.exists(path),
                      "config": cfg, "modes": MODE_HELP},
                     ensure_ascii=False, indent=2))
    return 0


def cmd_doctor(args) -> int:
    """自检：配置是否完整、入口网址是否可达、是否被踢到登录页。"""
    cfg = load_config(config_path(args))
    access = cfg["access"]
    login_type = access.get("login_type", "")
    entry = access.get("entry_url") or ""
    report = {
        "config": config_path(args),
        "登录方式": LOGIN_HELP.get(access.get("login_type", ""), {}).get("名称", "未设置"),
        "mode": access.get("mode"),
        "说明": MODE_HELP.get(access.get("mode", ""), ""),
        "entry_url": entry,
        "download_dir": cfg.get("download_dir"),
        "checks": [],
    }
    ok = True

    def add(name, passed, detail, fatal=True):
        nonlocal ok
        report["checks"].append({"项": name,
                                 "结果": "通过" if passed else ("失败" if fatal else "提示"),
                                 "详情": detail})
        if fatal:
            ok = ok and passed

    add("配置文件存在", os.path.exists(config_path(args)), config_path(args))
    add("入口网址已设置", bool(entry), entry or "未设置，先跑 init")
    add("下载目录已设置", bool(cfg.get("download_dir")), cfg.get("download_dir") or "未设置")
    if cfg.get("download_dir"):
        add("下载目录可写", os.access(cfg["download_dir"], os.W_OK),
            cfg["download_dir"])

    if entry:
        try:
            req = urllib.request.Request(
                entry, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
            with urllib.request.urlopen(req, timeout=args.timeout) as r:
                body = r.read(20000).decode("utf-8", "ignore")
                final = r.geturl()
            landed_login = bool(re.search(
                r"authserver|/login|统一身份认证|CAS|sso", final + body, re.I))
            add("入口网址可达", True, f"HTTP {r.status} -> {final[:120]}")
            if landed_login:
                hint = ("命令行没有浏览器的 cookie，看到登录页是正常的；"
                        + ("以浏览器里 https://www.cnki.net 能打开、右上角显示你的账号名为准，"
                           "不用管这一条"
                           if login_type in ("", "personal") else
                           "以浏览器里能打开该网址、且页头显示学校/单位名为准，不用管这一条"))
            else:
                hint = "命令行也直接进了知网，说明免登录就能访问"
            add("浏览器登录态", not landed_login, hint, fatal=False)
        except Exception as exc:                                   # noqa: BLE001
            add("入口网址可达", False, f"{type(exc).__name__}: {exc}")

    report["结论"] = "配置可用" if ok else "还有问题，按上面失败的项处理"
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if ok else 1


def cmd_proxy(args) -> int:
    """把已存的知网原生网址换成本校可访问的网址（可重复运行）。"""
    cfg = load_config(config_path(args))
    if getattr(args, "school", ""):
        args.school_domain = args.school
    if args.mode:
        cfg["access"]["mode"] = args.mode
    if args.school_domain:
        cfg["access"]["school_domain"] = args.school_domain
        cfg["school"]["domain"] = args.school_domain
    if args.prefix:
        cfg["access"]["prefix"] = args.prefix.rstrip("/")
    if args.ezproxy:
        cfg["access"]["ezproxy"] = args.ezproxy

    state, items = load_state(args.workdir)
    changed, unsupported = 0, []
    for it in items:
        if not it.get("url"):
            continue
        new, supported = rewrite_url(it["url"], cfg)
        if not supported:
            unsupported.append(it.get("n"))
            continue
        if new != it["url"]:
            it["url"] = new
            changed += 1
    if changed:
        save_json(os.path.join(args.workdir, "items.json"), items)
        sync_state(args.workdir)
    out = {
        "mode": cfg["access"]["mode"],
        "rewritten": changed,
        "remaining": sum(1 for i in items if not i.get("url")),
        "sample": next((it["url"] for it in items if it.get("url")), ""),
    }
    if unsupported:
        out["无法改写"] = unsupported
        out["hint"] = "这些条目缺少必要配置，先跑 doctor 看哪一项没配。"
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0 if not unsupported else 1


def cmd_parse(args) -> int:
    if not os.path.exists(args.input):
        print(json.dumps({"error": "清单文件不存在: " + args.input,
                          "hint": "把清单放到这个路径，或在参数里改成真实路径。"},
                         ensure_ascii=False))
        return 2
    cfg = load_config(config_path(args))
    if not args.download_dir and cfg.get("download_dir"):
        args.download_dir = cfg["download_dir"]
    items = parse_manifest(args.input)
    items_path, state_path = workdir_paths(args.workdir)
    old_state = load_json(state_path, {})
    old_items = old_state.get("items") or {}
    carried = 0
    for item in items:
        prev = old_items.get(item["key"])
        if not prev:
            continue
        for field in ("status", "file", "reason", "updated"):
            if prev.get(field):
                item[field] = prev[field]
                carried += 1
    save_json(items_path, items)
    save_json(state_path, {
        "items": old_items,
        "created": old_state.get("created") or time.strftime("%Y-%m-%d %H:%M:%S"),
        "download_dir": args.download_dir or old_state.get("download_dir", ""),
    })
    sync_state(args.workdir)
    kinds: dict[str, int] = {}
    needs_search = 0
    for item in items:
        kinds[item["kind"]] = kinds.get(item["kind"], 0) + 1
        if not item["url"]:
            needs_search += 1
    print(json.dumps({
        "parsed": len(items),
        "kinds": kinds,
        "needs_search": needs_search,
        "resumed": carried > 0,
        "items_path": items_path,
        "state_path": state_path,
        "items": items,
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_list(args) -> int:
    _, items = load_state(args.workdir)
    if not items:
        print(json.dumps({"error": "no items.json in " + args.workdir},
                         ensure_ascii=False))
        return 2
    only = args.status or ""
    rows = [i for i in items if not only or i.get("status", "pending") == only]
    print(json.dumps({"count": len(rows), "items": rows},
                     ensure_ascii=False, indent=2))
    return 0


def cmd_next(args) -> int:
    _, items = load_state(args.workdir)
    pending = [i for i in items if i.get("status", "pending") == "pending"]
    if not pending:
        print(json.dumps({"done": True, "pending": 0}, ensure_ascii=False))
        return 0
    print(json.dumps({"done": False, "pending": len(pending), "item": pending[0]},
                     ensure_ascii=False, indent=2))
    return 0


def cmd_mark(args) -> int:
    _, items = load_state(args.workdir)
    item = find_item(items, key=args.key, n=args.n, title=args.title)
    if item is None:
        print(json.dumps({"error": "item not found"}, ensure_ascii=False))
        return 2
    if args.status not in STATUSES:
        print(json.dumps({"error": "bad status: " + args.status},
                         ensure_ascii=False))
        return 2
    item["status"] = args.status
    if args.status == "pending":          # 重新排队时清掉上一轮的文件/原因
        item["file"] = ""
        item["reason"] = ""
    if args.file:
        item["file"] = args.file
    if args.reason is not None:
        item["reason"] = args.reason
    elif args.status in ("ok", "skipped"):
        item["reason"] = ""      # 成功/跳过时清掉旧的失败原因，否则报告自相矛盾
    if args.url:
        item["url"] = args.url
    item["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")
    save_json(workdir_paths(args.workdir)[0], items)
    sync_state(args.workdir)
    ok = sum(1 for i in items if i.get("status") == "ok")
    failed = sum(1 for i in items if i.get("status") == "failed")
    skipped = sum(1 for i in items if i.get("status") == "skipped")
    pending = sum(1 for i in items if i.get("status", "pending") == "pending")
    print(json.dumps({
        "marked": {"n": item.get("n"), "title": item.get("title"),
                   "status": item["status"]},
        "counts": {"ok": ok, "failed": failed, "skipped": skipped,
                   "pending": pending},
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_scan(args) -> int:
    since = args.since or 0.0
    found = []
    if os.path.isdir(args.dir):
        for name in sorted(os.listdir(args.dir)):
            full = os.path.join(args.dir, name)
            try:
                st = os.stat(full)
            except OSError:
                continue
            if not os.path.isfile(full) or st.st_mtime <= since:
                continue
            found.append({
                "name": name,
                "path": full,
                "size": st.st_size,
                "mtime": st.st_mtime,
                "mtime_local": time.strftime("%Y-%m-%d %H:%M:%S",
                                             time.localtime(st.st_mtime)),
            })
    print(json.dumps({"dir": args.dir, "since": since, "now": time.time(),
                      "count": len(found), "files": found},
                     ensure_ascii=False, indent=2))
    return 0


DOC_EXT = (".pdf", ".caj", ".kdh", ".nh")


def _new_docs(directory: str, since: float) -> list:
    """指定目录里 since 之后新出现的正文文件（含大小）。"""
    out = []
    if not os.path.isdir(directory):
        return out
    for name in sorted(os.listdir(directory)):
        full = os.path.join(directory, name)
        try:
            st = os.stat(full)
        except OSError:
            continue
        if not os.path.isfile(full) or st.st_mtime <= since:
            continue
        if os.path.splitext(name)[1].lower() not in DOC_EXT:
            continue
        out.append({"name": name, "path": full, "size": st.st_size,
                    "ok": st.st_size > 10240})
    return out


def cmd_waitnew(args) -> int:
    """等新文件落地。慢速下载（40 MB 学位论文）靠这个判定，不靠猜时间。"""
    started = time.time()
    deadline = started + max(1.0, args.timeout)
    found = []
    while True:
        found = _new_docs(args.dir, args.since or 0.0)
        ready = [f for f in found if f["ok"]]
        if ready:
            break
        if time.time() >= deadline:
            break
        time.sleep(max(0.2, args.interval))
    ready = [f for f in found if f["ok"]]
    print(json.dumps({
        "dir": args.dir,
        "since": args.since or 0.0,
        "waited_seconds": round(time.time() - started, 1),
        "count": len(ready),
        "files": ready,
        "small_or_partial": [f for f in found if not f["ok"]],
        "hint": "" if ready else "没有新文件：确认点击后没开新标签、没被付费墙拦；可重试点击一次",
    }, ensure_ascii=False, indent=2))
    return 0 if ready else 1


def cmd_existing(args) -> int:
    _, items = load_state(args.workdir)
    files = []
    if os.path.isdir(args.dir):
        for name in os.listdir(args.dir):
            if os.path.splitext(name)[1].lower() in (".pdf", ".caj", ".kdh", ".nh"):
                files.append((name, os.path.join(args.dir, name)))
    hits = []
    for item in items:
        if item.get("status") in ("ok",) and item.get("file"):
            continue
        best, score, best_len = None, 0.0, 10 ** 9
        for name, full in files:
            stem = os.path.splitext(name)[0]
            s = 0.0
            for tv in _title_variants(item["title"]):
                for sv in _stem_variants(stem):
                    s = max(s, similar(tv, sv))
            # 分数相同时优先取名字短的，避免选中「xxx (1).pdf」这种重复副本
            if s > score + 1e-6 or (abs(s - score) <= 1e-6 and len(name) < best_len):
                best, score, best_len = full, s, len(name)
        if best and score >= args.threshold:
            item["status"] = "skipped"
            item["file"] = best
            item["reason"] = "已存在同名文件"
            item["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")
            hits.append({"n": item.get("n"), "title": item["title"],
                         "file": best, "score": round(score, 3)})
    if hits:
        save_json(workdir_paths(args.workdir)[0], items)
        sync_state(args.workdir)
    print(json.dumps({"matched": len(hits), "hits": hits},
                     ensure_ascii=False, indent=2))
    return 0


def cmd_report(args) -> int:
    state, items = sync_state(args.workdir)
    buckets = {"ok": [], "failed": [], "skipped": [], "pending": []}
    for item in items:
        buckets[item.get("status", "pending")].append(item)

    lines = ["# CNKI 批量下载报告", ""]
    lines.append("- 生成时间：%s" % time.strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("- 清单：%d 篇" % len(items))
    lines.append("- 成功 %d / 跳过 %d / 失败 %d / 待处理 %d"
                 % (len(buckets["ok"]), len(buckets["skipped"]),
                    len(buckets["failed"]), len(buckets["pending"])))
    lines.append("")
    for name, label in (("ok", "已下载"), ("skipped", "已存在，跳过"),
                        ("failed", "失败"), ("pending", "待处理")):
        rows = buckets[name]
        lines.append("## %s（%d）" % (label, len(rows)))
        lines.append("")
        if not rows:
            lines.append("无")
            lines.append("")
            continue
        lines.append("| # | 题名 | 作者 | 年份 | 文件 | 原因 |")
        lines.append("|---|---|---|---|---|---|")
        for item in rows:
            lines.append("| %s | %s | %s | %s | %s | %s |" % (
                item.get("n", ""),
                (item.get("title") or "").replace("|", "\\|"),
                (item.get("authors") or "").replace("|", "\\|"),
                item.get("year", ""),
                (item.get("file") or "").replace("|", "\\|"),
                (item.get("reason") or "").replace("|", "\\|"),
            ))
        lines.append("")

    os.makedirs(args.workdir, exist_ok=True)
    report_path = os.path.join(args.workdir, "report.md")
    with io.open(report_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))

    failed_path = os.path.join(args.workdir, "failed.md")
    failed_lines = ["# 未下载成功的文献", ""]
    if not buckets["failed"]:
        failed_lines.append("无。")
    for item in buckets["failed"]:
        failed_lines.append("- [ ] %s（%s，%s）— %s"
                            % (item.get("title", ""), item.get("authors", ""),
                               item.get("year", ""), item.get("reason", "")))
    with io.open(failed_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(failed_lines) + "\n")

    print(json.dumps({
        "total": len(items),
        "ok": len(buckets["ok"]),
        "skipped": len(buckets["skipped"]),
        "failed": len(buckets["failed"]),
        "pending": len(buckets["pending"]),
        "report": report_path,
        "failed_list": failed_path,
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_sim(args) -> int:
    print(json.dumps({"score": round(similar(args.a, args.b), 3),
                      "norm_a": norm(args.a), "norm_b": norm(args.b)},
                     ensure_ascii=False))
    return 0


def main(argv: list[str]) -> int:
    # 必须在 parse_args 之前设定，否则 --help 与参数错误提示会按 GBK 输出成乱码
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except Exception:
                pass
    parser = argparse.ArgumentParser(description="CNKI batch download helper")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("parse", help="parse a manifest into items.json")
    p.add_argument("input")
    p.add_argument("--workdir", default="cnki-batch")
    p.add_argument("--download-dir", default="")
    p.set_defaults(func=cmd_parse)

    p = sub.add_parser("list", help="list items")
    p.add_argument("--workdir", default="cnki-batch")
    p.add_argument("--status", default="")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("next", help="print next pending item")
    p.add_argument("--workdir", default="cnki-batch")
    p.set_defaults(func=cmd_next)

    p = sub.add_parser("mark", help="set one item's status")
    p.add_argument("--workdir", default="cnki-batch")
    p.add_argument("--key", default="")
    p.add_argument("--n", type=int, default=0)
    p.add_argument("--title", default="")
    p.add_argument("--status", required=True)
    p.add_argument("--file", default="")
    p.add_argument("--reason", default=None)
    p.add_argument("--url", default="")
    p.set_defaults(func=cmd_mark)

    p = sub.add_parser("scan", help="files newer than a marker timestamp")
    p.add_argument("--dir", required=True)
    p.add_argument("--since", type=float, default=0.0)
    p.set_defaults(func=cmd_scan)

    p = sub.add_parser("waitnew", help="等下载目录出现新文件（默认最多 15 秒）")
    p.add_argument("--dir", required=True, help="浏览器下载保存目录，例如 C:\\Users\\你的名字\\Downloads")
    p.add_argument("--since", type=float, default=0.0, help="基准时间戳（epoch 秒），只看比它新的文件")
    p.add_argument("--timeout", type=float, default=15.0, help="最多等多少秒（大文件放到 30）")
    p.add_argument("--interval", type=float, default=1.0, help="每隔多少秒看一次")
    p.set_defaults(func=cmd_waitnew)

    p = sub.add_parser("existing", help="match pending items to local files")
    p.add_argument("--workdir", default="cnki-batch")
    p.add_argument("--dir", required=True)
    p.add_argument("--threshold", type=float, default=0.82)
    p.set_defaults(func=cmd_existing)

    p = sub.add_parser("report", help="write report.md and failed.md")
    p.add_argument("--workdir", default="cnki-batch")
    p.set_defaults(func=cmd_report)

    p = sub.add_parser("init", help="一次性配置：先定登录方式，再定网址")
    p.add_argument("--config", default="")
    p.add_argument("--login", default="",
                   help="personal|institution|webvpn|prefix_vpn|ezproxy")
    p.add_argument("--from-url", default="", help="浏览器里能打开的那个知网网址")
    p.add_argument("--mode", default=None, choices=list(MODE_HELP))
    p.add_argument("--school-name", default="", help="学校名字，如 示例大学")
    p.add_argument("--school-domain", default="", help="学校域名，如 xx.edu.cn")
    p.add_argument("--prefix", default="", help="VPN 路径前缀")
    p.add_argument("--ezproxy", default="", help="图书馆代理前缀")
    p.add_argument("--entry-url", default="", help="登录后能打开知网的入口网址")
    p.add_argument("--download-dir", default="")
    p.set_defaults(func=cmd_init)

    p = sub.add_parser("config", help="查看当前配置")
    p.add_argument("--config", default="")
    p.set_defaults(func=cmd_config)

    p = sub.add_parser("doctor", help="自检：配置、入口网址、下载目录")
    p.add_argument("--config", default="")
    p.add_argument("--timeout", type=float, default=15.0)
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("proxy", help="把已存网址换成本校可访问的网址")
    p.add_argument("--config", default="")
    p.add_argument("--workdir", default="cnki-batch")
    p.add_argument("--mode", default=None, choices=list(MODE_HELP))
    p.add_argument("--school-domain", default="")
    p.add_argument("--prefix", default="")
    p.add_argument("--ezproxy", default="")
    p.add_argument("--school", default="", help="（旧参数）等同 --school-domain")
    p.set_defaults(func=cmd_proxy)

    p = sub.add_parser("sim", help="title similarity score")
    p.add_argument("a")
    p.add_argument("b")
    p.set_defaults(func=cmd_sim)

    args = parser.parse_args(argv)
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
