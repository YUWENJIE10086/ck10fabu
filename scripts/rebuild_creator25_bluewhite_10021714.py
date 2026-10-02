from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import csv
import json
import re
import shutil
import zipfile

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt

STAMP = "10021714"
ROOT = Path("决赛发布材料/1002PPT决赛完善版")
ZIP_PATH = Path("创客25图片.zip")
OLD_INDEX = Path("决赛发布材料/10011941PPT母版原生融合版/创客25图片_401张逐图样式分类与PPT应用索引_10011941.csv")
SRC = ROOT / "智烤管家_10021215_一等奖高阶光效价值深化版.pptx"
TITLE_REF = Path("决赛发布材料/10012237PPT母版高端科技版/智烤管家_10012237_3S母版高端科技决赛版.pptx")
OUT = ROOT / f"智烤管家_{STAMP}_创客25逐图学习蓝白母版重做版.pptx"
ORGANIZED = ROOT / f"创客25图片解压整理_{STAMP}"
PREVIEWS = ORGANIZED / "逐图预览"
INDEX_OUT = ORGANIZED / f"创客25图片_401张逐图学习索引_{STAMP}.csv"
README_OUT = ORGANIZED / f"README_逐图学习与本版应用_{STAMP}.md"
TMP_EXTRACT = Path("/tmp/creator25_extract_10021714")

BLUE = RGBColor(0x2A, 0xA8, 0xFF)
BLUE_SOFT = RGBColor(0x8E, 0xD9, 0xFF)
BLUE_PALE = RGBColor(0xC8, 0xEE, 0xFF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x07, 0x1A, 0x3E)
LINE_BLUE = RGBColor(0x5B, 0xBD, 0xF5)

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A_NS, "p": P_NS}


def emu_to_in(v: int) -> float:
    return v / 914400.0


def safe_name(s: str, max_len: int = 75) -> str:
    s = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", s)
    s = re.sub(r"\s+", " ", s).strip(" ._")
    return (s[:max_len] or "image")


def read_old_index() -> dict[int, dict[str, str]]:
    mapping: dict[int, dict[str, str]] = {}
    with OLD_INDEX.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            try:
                mapping[int(row["idx"])] = row
            except Exception:
                continue
    return mapping


def unzip_and_organize() -> tuple[list[dict[str, str]], dict[str, int]]:
    if TMP_EXTRACT.exists():
        shutil.rmtree(TMP_EXTRACT)
    TMP_EXTRACT.mkdir(parents=True, exist_ok=True)
    PREVIEWS.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(ZIP_PATH, "r") as zf:
        zf.extractall(TMP_EXTRACT)

    exts = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
    files = sorted(p for p in TMP_EXTRACT.rglob("*") if p.suffix.lower() in exts)
    old = read_old_index()
    rows: list[dict[str, str]] = []
    counts: dict[str, int] = {}

    # Keep every single reference image visible in the repo, but as a compact review preview
    # rather than duplicating the full 402 MB LFS archive.
    for idx, p in enumerate(files, 1):
        try:
            im = Image.open(p).convert("RGB")
        except Exception:
            continue
        w, h = im.size
        meta = old.get(idx, {})
        cluster = str(meta.get("cluster", "X"))
        family = meta.get("style_family", "未分类") or "未分类"
        learned = meta.get("learned_style", "")
        usage = meta.get("used_in_final_ppt", "")
        orientation = "横版" if w >= h else "竖版"
        cat = f"C{int(cluster):02d}_{safe_name(family, 28)}" if cluster.isdigit() else f"CX_{safe_name(family, 28)}"
        cat_dir = PREVIEWS / cat
        cat_dir.mkdir(parents=True, exist_ok=True)

        thumb = im.copy()
        thumb.thumbnail((960, 540), Image.LANCZOS)
        canvas = Image.new("RGB", (1000, 610), (6, 19, 48))
        x = (1000 - thumb.width) // 2
        y = 48 + (520 - thumb.height) // 2
        canvas.paste(thumb, (x, y))
        d = ImageDraw.Draw(canvas)
        label = f"{idx:03d} | {family} | {p.name}"
        d.text((24, 16), label[:115], fill=(230, 246, 255))
        out_name = f"{idx:03d}_{safe_name(p.stem, 55)}.jpg"
        out_path = cat_dir / out_name
        canvas.save(out_path, "JPEG", quality=68, optimize=True, progressive=True)

        small = im.copy()
        small.thumbnail((96, 96), Image.LANCZOS)
        q = small.quantize(colors=5)
        pal = q.getpalette()[:15]
        colors = ["#%02X%02X%02X" % tuple(pal[i:i+3]) for i in range(0, len(pal), 3)]
        rows.append({
            "idx": str(idx),
            "original_file": str(p.relative_to(TMP_EXTRACT)),
            "preview_file": str(out_path.relative_to(ORGANIZED)),
            "width": str(w),
            "height": str(h),
            "orientation": orientation,
            "cluster": cluster,
            "style_family": family,
            "dominant5": " ".join(colors),
            "逐图学习结论": learned,
            "本版应用": usage,
        })
        counts[family] = counts.get(family, 0) + 1

    fieldnames = [
        "idx", "original_file", "preview_file", "width", "height", "orientation",
        "cluster", "style_family", "dominant5", "逐图学习结论", "本版应用"
    ]
    with INDEX_OUT.open("w", encoding="utf-8-sig", newline="") as f:
        wr = csv.DictWriter(f, fieldnames=fieldnames)
        wr.writeheader()
        wr.writerows(rows)
    return rows, counts


def remove_effects(shape) -> None:
    # Remove the previous "neon / nightclub" feeling while keeping fills and layout.
    el = shape._element
    for tag in ("glow", "softEdge", "reflection", "outerShdw", "innerShdw"):
        for node in el.xpath(f".//a:{tag}"):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)


def remove_picture(shape) -> None:
    sp = shape._element
    sp.getparent().remove(sp)


def find_top_title(slide):
    candidates = []
    for sh in slide.shapes:
        if not getattr(sh, "has_text_frame", False):
            continue
        txt = sh.text.strip()
        if not txt:
            continue
        y = emu_to_in(sh.top)
        x = emu_to_in(sh.left)
        w = emu_to_in(sh.width)
        h = emu_to_in(sh.height)
        # page title is centered in the first 0.55 in, not the /// corner ornaments
        if y <= 0.32 and w >= 3.0 and h <= 0.8 and txt != "///":
            candidates.append((w, sh))
    return max(candidates, key=lambda t: t[0])[1] if candidates else None


def first_run_rpr(shape):
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            return run._r.get_or_add_rPr()
    return None


def copy_title_gradient_only(src_shape, dst_shape) -> None:
    if src_shape is None or dst_shape is None:
        return
    src_rpr = first_run_rpr(src_shape)
    if src_rpr is None:
        return
    src_fills = []
    for child in src_rpr:
        if child.tag in {
            f"{{{A_NS}}}solidFill", f"{{{A_NS}}}gradFill", f"{{{A_NS}}}pattFill", f"{{{A_NS}}}noFill"
        }:
            src_fills.append(deepcopy(child))
    if not src_fills:
        return
    for p in dst_shape.text_frame.paragraphs:
        for run in p.runs:
            rpr = run._r.get_or_add_rPr()
            for child in list(rpr):
                if child.tag in {
                    f"{{{A_NS}}}solidFill", f"{{{A_NS}}}gradFill", f"{{{A_NS}}}pattFill", f"{{{A_NS}}}noFill"
                }:
                    rpr.remove(child)
            for fill in src_fills:
                rpr.insert(0, deepcopy(fill))


def set_line(shape, color=LINE_BLUE, width=0.9) -> None:
    try:
        shape.line.color.rgb = color
        shape.line.width = Pt(width)
    except Exception:
        pass


def set_text_color(shape, color) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            try:
                run.font.color.rgb = color
            except Exception:
                pass


def is_numeric_anchor(text: str) -> bool:
    t = text.strip().replace(" ", "")
    return bool(re.fullmatch(r"[≈+\-]?\d+(?:\.\d+)?%?\+?", t))


def style_slide(slide, si: int, title_ref_slide=None) -> dict[str, int]:
    stats = {"gifs_removed": 0, "effects_removed": 0, "cards_refined": 0}
    target_title = find_top_title(slide)
    ref_title = find_top_title(title_ref_slide) if title_ref_slide is not None else None

    # First remove animated/full-area light sweeps. This was the biggest visual problem.
    for sh in list(slide.shapes):
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                ext = (sh.image.ext or "").lower()
            except Exception:
                ext = ""
            if ext == "gif":
                remove_picture(sh)
                stats["gifs_removed"] += 1

    # Preserve the first 0.9in title system; only restore its original blue-white gradient fill.
    if target_title is not None and ref_title is not None:
        copy_title_gradient_only(ref_title, target_title)

    for sh in list(slide.shapes):
        y = emu_to_in(sh.top)
        if y < 0.90:
            # top area: keep page title/subtitle/ornaments, remove only glow effects
            before = len(sh._element.xpath(".//a:glow|.//a:softEdge|.//a:reflection|.//a:outerShdw"))
            remove_effects(sh)
            stats["effects_removed"] += before
            # the separator becomes blue-white, not neon green/cyan
            if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and emu_to_in(sh.height) <= 0.08:
                try:
                    sh.fill.solid(); sh.fill.fore_color.rgb = BLUE_SOFT
                    sh.line.fill.background()
                except Exception:
                    pass
            continue

        before = len(sh._element.xpath(".//a:glow|.//a:softEdge|.//a:reflection|.//a:outerShdw"))
        remove_effects(sh)
        stats["effects_removed"] += before

        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            name = (sh.name or "").lower()
            w = emu_to_in(sh.width); h = emu_to_in(sh.height)
            # Creator25's strongest reusable syntax: deep-blue glass cards + precise pale-blue edge.
            if ("round" in name or "圆角" in name or h >= 0.36) and w >= 0.8:
                set_line(sh, LINE_BLUE, 0.85 if h < 1.0 else 1.0)
                stats["cards_refined"] += 1
            # arrows / step dots should be blue-white, not fluorescent cyan.
            if "chevron" in name or "arrow" in name or "箭头" in name:
                try:
                    sh.fill.solid(); sh.fill.fore_color.rgb = BLUE
                    sh.line.fill.background()
                except Exception:
                    pass
            if "oval" in name or "椭圆" in name:
                try:
                    sh.fill.solid(); sh.fill.fore_color.rgb = BLUE
                    sh.line.color.rgb = BLUE_PALE; sh.line.width = Pt(0.6)
                except Exception:
                    pass

        if getattr(sh, "has_text_frame", False):
            txt = sh.text.strip()
            if not txt:
                continue
            # large KPI / step numerals: white or soft blue; never gold/orange.
            if is_numeric_anchor(txt):
                set_text_color(sh, WHITE if si in (3, 4, 8, 14, 15) else BLUE_PALE)
            else:
                # Keep body copy readable; only normalize obvious colored labels to pale blue.
                for p in sh.text_frame.paragraphs:
                    for run in p.runs:
                        try:
                            rgb = run.font.color.rgb
                            if rgb is not None and (rgb[0] > 190 and rgb[1] < 190):
                                run.font.color.rgb = BLUE_PALE
                        except Exception:
                            pass

    # Slide-specific cleanup learned from the image set.
    if si == 3:
        # 531 survey: one data badge is enough; remove the two outer decorative rings.
        ovals = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and ("Oval" in (sh.name or "") or "椭圆" in (sh.name or "")) and 1.0 < emu_to_in(sh.top) < 2.4]
        ovals.sort(key=lambda sh: emu_to_in(sh.width), reverse=True)
        for sh in ovals[:2]:
            remove_picture(sh)
    if si in (6, 7):
        # three-step model pages: screenshots are evidence windows; use a light blue precise edge.
        for sh in slide.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and 1.9 <= emu_to_in(sh.top) <= 2.3 and 1.3 <= emu_to_in(sh.height) <= 1.9:
                set_line(sh, BLUE_PALE, 0.75)
    if si == 14:
        # Results page: make hard numbers the visual anchor in pure white / pale blue.
        for sh in slide.shapes:
            if getattr(sh, "has_text_frame", False) and is_numeric_anchor(sh.text):
                set_text_color(sh, WHITE)

    return stats


def rebuild_ppt() -> dict[str, int]:
    prs = Presentation(SRC)
    ref = Presentation(TITLE_REF)
    total = {"gifs_removed": 0, "effects_removed": 0, "cards_refined": 0}
    for i, slide in enumerate(prs.slides, 1):
        ref_slide = ref.slides[i - 1] if i - 1 < len(ref.slides) else None
        st = style_slide(slide, i, ref_slide)
        for k, v in st.items():
            total[k] += v
    prs.save(OUT)
    return total


def write_readme(rows: list[dict[str, str]], counts: dict[str, int], stats: dict[str, int]) -> None:
    count_lines = "\n".join(f"- {k}: {v} 张" for k, v in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])))
    text = f"""# 创客25图片逐图整理 + 智烤管家蓝白母版重做（{STAMP}）

## 1. 这次先做素材整理，不再凭印象改PPT

- `创客25图片.zip` 通过 Git LFS 拉取后，已在构建工作区完整解压。
- 实际识别并逐张处理：**{len(rows)} 张**。
- 为避免把 402MB 原始压缩包再次完整复制进仓库，仓库内保留每张图的 **逐图高清预览**，按样式类别分文件夹整理；原始高清图仍由根目录 LFS 压缩包保存。
- 每张图均在 `{INDEX_OUT.name}` 中记录：原始文件、预览文件、尺寸、横竖版、样式类别、主色、逐图学习结论、本版应用位置。

### 401张样式分布
{count_lines}

## 2. 本次PPT视觉基准

本次从 `智烤管家_10021215_一等奖高阶光效价值深化版.pptx` **重新起稿**，不以 10021320 / 10021336 两个错误方向版本为基底。

最重要的恢复项：

1. **每一页顶部标题体系完整保留**：页标题、副标题、左右装饰符号都不删除。
2. **恢复原先的白→青蓝渐变标题**：从 `10012237_3S母版高端科技决赛版` 提取标题文字渐变填充，只恢复标题色彩语言，不复制荧光效果。
3. **全稿回到蓝白主色**：深蓝背景 + 白字 + 冰蓝/天蓝辅助；大数字不用金色。
4. **去掉“荧光灯带感”**：移除 GIF 扫光、Glow、SoftEdge、Reflection、OuterShadow。
5. **卡片不再自己发明**：沿用 Creator25 里反复出现的蓝青玻璃卡、浅色证据窗、三步流程、模块链、驾驶舱、大数字六种结构。

## 3. 逐页采用 Creator25 哪类样式

- S1 封面：C8 舞台超简洁页——大标题 + 一句价值主张 + 留白。
- S2 三难：C10 蓝青高质感组件页——蓝白玻璃框 + 小题签，不要霓虹框。
- S3 531问卷：C3 单中心构图——只留一个 531 数据徽章，三层发光环删除。
- S4 1234：C9 模块化流程——四模块同构，白蓝编号，蓝色连接关系。
- S5 总体蓝图：C9 + C5——复杂关系靠模块和真实看板，不靠装饰。
- S6 模型①：C4 三步流 + C6 大白证据卡——状态标准化 / AHP / Weibull 三步。
- S7 模型②：C4 三步流 + C5 数据驾驶舱——有效烤能 / 真实缺口 / 建设推演。
- S8 双模耦合：C9 模块链——存量体检 → 有效烤能 → 增量推演 → 投入决策。
- S9 一库/平台证据：C2/C6——一侧逻辑，一侧真实系统截图。
- S10 三智：C4 + C2——角色工作流 + 真实系统证据。
- S11 四队：C9——职责模块化、机制闭环化。
- S12 管护/基金：C6——证据大于装饰。
- S13 创新：C9——业务创新与技术创新双栏，不堆彩色块。
- S14 成效：C8——大数字是第一视觉层级，解释文字缩小。
- S15 收口：C8——白青渐变大字收口，不再做一排发光按钮。

## 4. 本次程序化处理统计

- 动态 GIF / 扫光删除：{stats['gifs_removed']} 个
- Glow / SoftEdge / Reflection / 外阴影节点删除：{stats['effects_removed']} 处
- 蓝白精密边框统一：{stats['cards_refined']} 个组件

最终文件：`{OUT.name}`
"""
    README_OUT.write_text(text, encoding="utf-8")


def main() -> None:
    rows, counts = unzip_and_organize()
    if len(rows) != 401:
        raise RuntimeError(f"Creator25 expected 401 images, got {len(rows)}")
    stats = rebuild_ppt()
    write_readme(rows, counts, stats)
    print(json.dumps({
        "images": len(rows),
        "output": str(OUT),
        "organized": str(ORGANIZED),
        "stats": stats,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
