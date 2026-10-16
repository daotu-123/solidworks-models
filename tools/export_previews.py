#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批量导出 SolidWorks 预览图 —— 供本仓库更新时使用。

GitHub 无法预览 .SLDPRT / .SLDASM / .SLDDRW，所以每个项目目录下都放了
preview/ 预览图。新增或修改模型后，运行本脚本重新生成即可。

环境要求
--------
* Windows + 已安装 SolidWorks（脚本通过 COM 调用，需本机授权）
* Python 3.9+，依赖 pywin32：  pip install pywin32

用法
----
# 导出某个项目（默认输出到 <项目>/preview/）
python tools/export_previews.py "四足爬行机器人"

# 指定输出目录 / 指定视图 / 指定宽度
python tools/export_previews.py "curved-surface" --out preview --view isometric --width 1000

# 只处理部分文件
python tools/export_previews.py "四足爬行机器人" --only 零件1 零件15

说明
----
* 输出文件名 = 相对路径把分隔符换成 '__'，例如
  quminajichu/杯子.SLDPRT -> quminajichu__杯子.SLDPRT.png
  重命名时可据此对照。
* 脚本会跳过「无实体且无曲面体」的空文件，并打印报告。
* SolidWorks COM 必须串行调用，请勿并行运行多个实例。
"""

from __future__ import annotations

import argparse
import os
import sys
import time

try:
    import pythoncom
    import win32com.client
    from win32com.client import VARIANT
except ImportError:
    sys.exit("缺少 pywin32，请先执行：pip install pywin32")

# ---- SolidWorks 常量 ----
DOC_PART, DOC_ASSEMBLY, DOC_DRAWING = 1, 2, 3
EXT_TO_DOCTYPE = {
    ".sldprt": DOC_PART,
    ".sldasm": DOC_ASSEMBLY,
    ".slddrw": DOC_DRAWING,
}

VIEWS = {
    "isometric": ("*Isometric", 7),
    "front":     ("*Front",     1),
    "back":      ("*Back",      2),
    "left":      ("*Left",      4),
    "right":     ("*Right",     3),
    "top":       ("*Top",       5),
    "bottom":    ("*Bottom",    6),
}

SOLID_BODY, SURFACE_BODY = 0, 1


def byref_int(value: int = 0):
    """构造 SolidWorks 需要的 by-ref int VARIANT（错误码/警告码出参）。"""
    return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, value)


def body_counts(model):
    """返回 (实体数, 曲面体数)；装配体/工程图返回 (None, None)。"""
    try:
        solids = model.GetBodies2(SOLID_BODY, True)
        surfaces = model.GetBodies2(SURFACE_BODY, True)
    except Exception:
        return None, None
    return (0 if solids is None else len(solids),
            0 if surfaces is None else len(surfaces))


def shrink(path: str, max_width: int) -> str:
    """把导出的 PNG 缩到 max_width 宽，返回说明文字。

    SolidWorks 只能按窗口尺寸导出，图片经常偏大；这里用 Pillow 等比缩小，
    以减小仓库体积。Pillow 缺失时原样保留，并在报告里给出提示。
    """
    try:
        from PIL import Image
    except ImportError:
        return "（未安装 Pillow，保留原始尺寸；pip install Pillow 可自动压缩）"

    with Image.open(path) as im:
        if im.width <= max_width:
            return f"{im.width}x{im.height}"
        height = round(im.height * max_width / im.width)
        im.resize((max_width, height), Image.LANCZOS).save(path, "PNG", optimize=True)
        return f"{max_width}x{height}（原 {im.width}x{im.height}）"


def collect(root: str, only: list[str] | None):
    """递归收集目标 CAD 文件，返回 [(绝对路径, 仓库内相对路径), ...]。"""
    targets = []
    for dirpath, _dirnames, filenames in os.walk(root):
        for name in sorted(filenames):
            if os.path.splitext(name)[1].lower() not in EXT_TO_DOCTYPE:
                continue
            full = os.path.join(dirpath, name)
            rel = os.path.relpath(full, root)
            if only and not any(k.lower() in rel.lower() for k in only):
                continue
            targets.append((full, rel))
    return targets


def main() -> int:
    ap = argparse.ArgumentParser(description="批量导出 SolidWorks 预览图")
    ap.add_argument("project", help="项目目录名（相对仓库根目录），如 四足爬行机器人")
    ap.add_argument("--repo", default=None, help="仓库根目录，默认取脚本上级目录")
    ap.add_argument("--out", default="preview", help="输出子目录名（默认 preview）")
    ap.add_argument("--view", default="isometric", choices=sorted(VIEWS),
                    help="视图方向（默认 isometric）")
    ap.add_argument("--width", type=int, default=1000, help="输出图片最大宽度（默认 1000）")
    ap.add_argument("--only", nargs="*", default=None, help="只处理路径包含这些关键字的文件")
    ap.add_argument("--keep-blank", action="store_true", help="同时导出空文件（默认跳过）")
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    repo = os.path.abspath(args.repo or os.path.dirname(here))
    root = os.path.join(repo, args.project)
    if not os.path.isdir(root):
        print(f"[x] 项目目录不存在：{root}")
        return 2

    out_dir = os.path.join(root, args.out)
    os.makedirs(out_dir, exist_ok=True)

    targets = collect(root, args.only)
    if not targets:
        print("[x] 未找到任何 CAD 文件")
        return 2

    view_name, view_id = VIEWS[args.view]
    print(f"仓库根目录 : {repo}")
    print(f"项目目录   : {root}")
    print(f"输出目录   : {out_dir}")
    print(f"视图 / 宽度: {args.view} / {args.width}px")
    print(f"待处理文件 : {len(targets)}\n")

    sw = win32com.client.Dispatch("SldWorks.Application")
    print(f"SolidWorks 已连接，版本 revision={sw.RevisionNumber}\n")

    ok = skipped = failed = 0
    for full, rel in targets:
        ext = os.path.splitext(full)[1].lower()
        doc_type = EXT_TO_DOCTYPE[ext]
        model = None
        try:
            model = sw.OpenDoc6(full, doc_type, 1, "",  byref_int(), byref_int())
            if model is None:
                print(f"  [x] 打开失败        {rel}")
                failed += 1
                continue

            title = model.GetTitle
            sw.ActivateDoc3(title, False, 0, byref_int())
            model = sw.ActiveDoc

            solids, surfaces = body_counts(model)
            if not args.keep_blank and (solids, surfaces) == (0, 0):
                print(f"  [-] 跳过空文件      {rel}")
                skipped += 1
                continue

            model.ForceRebuild3(False)
            try:
                model.ShowNamedView2(view_name, view_id)
            except Exception:
                pass
            model.ViewZoomtofit2()
            model.ClearSelection2(True)

            safe = rel.replace("\\", "__").replace("/", "__")
            dst = os.path.join(out_dir, safe + ".png")
            if os.path.exists(dst):
                os.remove(dst)

            # 注意：SolidWorks 用 SaveAs() 导出图片；SaveAs3() 对图像格式会报
            # "无效的参数数目" (0x8002000E)，不要改用它。
            if model.SaveAs(dst) and os.path.exists(dst):
                dim = shrink(dst, args.width)
                kb = os.path.getsize(dst) / 1024
                extra = f"实体={solids} 曲面={surfaces}" if solids is not None else ""
                print(f"  [+] {kb:8.1f} KB  {dim:<24} {rel}   {extra}")
                ok += 1
            else:
                print(f"  [x] 保存失败        {rel}")
                failed += 1

        except Exception as exc:  # noqa: BLE001
            print(f"  [x] 异常            {rel} -- {exc}")
            failed += 1
        finally:
            if model is not None:
                try:
                    sw.CloseDoc(model.GetTitle)
                except Exception:
                    pass

    print(f"\n完成：成功 {ok}，跳过 {skipped}，失败 {failed}")
    if skipped:
        print("提示：跳过的空文件只有基准面/原点，可考虑直接从仓库删除。")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
