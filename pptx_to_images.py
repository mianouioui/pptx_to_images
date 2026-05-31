#!/usr/bin/env python3
from __future__ import annotations

import argparse
import contextlib
import itertools
import os
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from pathlib import Path

__version__ = "1.1.0"

SUPPORTED_SUFFIXES = {".pptx", ".ppt"}
TRAILING_NUMBER_RE = re.compile(r"(\d+)(?=\.png$)", re.IGNORECASE)


@dataclass(frozen=True)
class EnginePlan:
    name: str
    detail: str


class ConversionError(RuntimeError):
    pass


INTERACTIVE_BANNER = """\
============================================================
  PPTX 转图片   v{version}
============================================================
  把 PPTX/PPT 演示文稿的每一页导出为 PNG 图片。

  怎么用：
    1. 把一个 PPTX 文件拖到这个窗口里（拖文件夹也行）
    2. 按回车

  输出：
    · 默认生成到同目录的 “文件名_images” 文件夹
    · 图片命名为 page001.png、page002.png ...
    · 拖入文件夹时，会转换里面的 PPTX/PPT 文件
    · 退出可按 Ctrl+C
------------------------------------------------------------"""


def parse_dropped_path(line: str) -> str | None:
    line = line.strip()
    if not line:
        return None
    if os.name == "nt":
        quote = line[0]
        if quote in {'"', "'"}:
            end = line.find(quote, 1)
            if end > 0:
                return line[1:end]
        parts = line.split()
        return parts[0] if parts else None
    try:
        parts = shlex.split(line, posix=True)
    except ValueError:
        return line
    return parts[0] if parts else None


def iter_presentations(folder: Path, recursive: bool) -> list[Path]:
    walker = folder.rglob("*") if recursive else folder.iterdir()
    candidates = [
        p for p in walker if p.suffix.lower() in SUPPORTED_SUFFIXES and p.is_file()
    ]
    return sorted(
        (p.resolve() for p in candidates),
        key=lambda p: p.name.lower(),
    )


def resolve_input_presentations(path: Path, recursive: bool) -> list[Path]:
    candidate = path.expanduser().resolve()
    if not candidate.exists():
        raise FileNotFoundError(str(candidate))
    if candidate.is_file():
        if candidate.suffix.lower() not in SUPPORTED_SUFFIXES:
            raise ConversionError(f"输入文件不是 PPTX/PPT：{candidate}")
        return [candidate]
    if candidate.is_dir():
        return iter_presentations(candidate, recursive)
    raise ConversionError(f"输入路径不是文件或文件夹：{candidate}")


def prompt_for_input(recursive: bool) -> list[Path] | None:
    print(INTERACTIVE_BANNER.format(version=__version__))
    while True:
        try:
            line = input("把 PPTX 文件拖到这里，然后按回车：")
        except (EOFError, KeyboardInterrupt):
            print()
            return None
        raw = parse_dropped_path(line)
        if raw is None:
            print("  还没有收到路径，请拖入一个 PPTX 文件或文件夹。\n")
            continue
        try:
            presentations = resolve_input_presentations(Path(raw), recursive)
        except (OSError, ConversionError) as exc:
            print(f"  读不到这个输入：{raw}\n  {exc}\n")
            continue
        if not presentations:
            print(f"  这个文件夹里没有 PPTX/PPT 文件：\n     {Path(raw).expanduser()}\n")
            continue
        return presentations


def unique_dir(path: Path, force: bool) -> Path:
    path = path.expanduser().resolve()
    if force:
        path.mkdir(parents=True, exist_ok=True)
        return path
    if not path.exists():
        path.mkdir(parents=True, exist_ok=False)
        return path
    for index in range(2, 1000):
        candidate = path.with_name(f"{path.name}_{index}")
        if not candidate.exists():
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate
    raise ConversionError(f"输出文件夹已存在且无法生成新名称：{path}")


def output_dir_for(presentation: Path, requested: Path | None, multiple: bool, force: bool) -> Path:
    if requested is None:
        return unique_dir(presentation.parent / f"{presentation.stem}_images", force)

    root = requested.expanduser()
    if multiple:
        return unique_dir(root / f"{presentation.stem}_images", force)
    return unique_dir(root, force)


def find_executable(names: list[str], extra_paths: list[Path] | None = None) -> str | None:
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    for path in extra_paths or []:
        if path.exists() and os.access(path, os.X_OK):
            return str(path)
    return None


def find_soffice() -> str | None:
    extra: list[Path] = []
    if sys.platform == "darwin":
        extra.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    if os.name == "nt":
        for base in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)")):
            if base:
                extra.append(Path(base) / "LibreOffice" / "program" / "soffice.exe")
    return find_executable(["soffice", "libreoffice"], extra)


def find_pdftoppm() -> str | None:
    return find_executable(["pdftoppm"])


SPINNER_FRAMES = (
    ("|", "/", "-", "\\")
    if os.name == "nt"
    else ("⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏")
)


@contextlib.contextmanager
def _progress(label: str):
    """在外部命令执行期间显示「正在工作」的进度，避免长任务看起来像卡死。"""
    stream = sys.stdout
    start = time.monotonic()

    if not stream.isatty():
        # 非交互终端（重定向 / 管道）：不画动画，只打印开始与结束，避免污染输出
        print(f"  · {label} …", flush=True)
        ok = False
        try:
            yield
            ok = True
        finally:
            elapsed = time.monotonic() - start
            print(f"  · {label} {'完成' if ok else '失败'}（{elapsed:.1f}s）", flush=True)
        return

    stop = threading.Event()

    def spin() -> None:
        for frame in itertools.cycle(SPINNER_FRAMES):
            if stop.is_set():
                break
            elapsed = time.monotonic() - start
            stream.write(f"\r  {frame} {label} …（{elapsed:.0f}s）   ")
            stream.flush()
            time.sleep(0.1)

    worker = threading.Thread(target=spin, daemon=True)
    worker.start()
    ok = False
    try:
        yield
        ok = True
    finally:
        stop.set()
        worker.join(timeout=1.0)
        elapsed = time.monotonic() - start
        mark = "✓" if ok else "✗"
        # 用空格覆盖残留的动画字符，再换行
        stream.write(f"\r  {mark} {label}（{elapsed:.1f}s）".ljust(64) + "\n")
        stream.flush()


def run_command(command: list[str], description: str, timeout: int = 180) -> subprocess.CompletedProcess[str]:
    with _progress(description):
        try:
            result = subprocess.run(
                command,
                check=False,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=timeout,
            )
        except FileNotFoundError as exc:
            raise ConversionError(f"找不到命令：{command[0]}") from exc
        except subprocess.TimeoutExpired as exc:
            raise ConversionError(f"{description} 超时（超过 {timeout} 秒未完成）。") from exc
        if result.returncode != 0:
            details = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
            raise ConversionError(f"{description} 失败。\n{details}".rstrip())
    return result


def convert_to_pdf_with_libreoffice(presentation: Path, work_dir: Path, soffice: str) -> Path:
    out_dir = work_dir / "pdf"
    out_dir.mkdir(parents=True, exist_ok=True)
    profile = work_dir / "lo-profile"
    profile.mkdir(parents=True, exist_ok=True)
    command = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--nodefault",
        "--nolockcheck",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(presentation),
    ]
    run_command(command, "LibreOffice 转 PDF", timeout=240)

    expected = out_dir / f"{presentation.stem}.pdf"
    if expected.exists() and expected.stat().st_size > 0:
        return expected

    pdfs = sorted(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if pdfs and pdfs[0].stat().st_size > 0:
        return pdfs[0]
    raise ConversionError("LibreOffice 没有生成可用的 PDF。")


def page_number_from_png(path: Path) -> int:
    match = TRAILING_NUMBER_RE.search(path.name)
    if match:
        return int(match.group(1))
    return 0


def move_numbered_pngs(raw_dir: Path, output_dir: Path, prefix: str, force: bool) -> int:
    pngs = sorted(
        (p for p in raw_dir.iterdir() if p.is_file() and p.suffix.lower() == ".png"),
        key=lambda p: (page_number_from_png(p), p.name.lower()),
    )
    if not pngs:
        raise ConversionError("没有生成 PNG 图片。")

    if force:
        # 复用已有输出目录时，先清掉上一次遗留的页图，
        # 避免源演示文稿页数变少时残留多余的尾页。
        for stale in output_dir.iterdir():
            if stale.is_file() and stale.suffix.lower() == ".png" and stale.name.startswith(prefix):
                stale.unlink()

    for index, source in enumerate(pngs, start=1):
        target = output_dir / f"{prefix}{index:03d}.png"
        if target.exists():
            if not force:
                raise ConversionError(f"输出图片已存在：{target}")
            target.unlink()
        shutil.move(str(source), str(target))
    return len(pngs)


def render_pdf_with_pdftoppm(
    pdf: Path,
    output_dir: Path,
    pdftoppm: str,
    width: int,
    prefix: str,
    force: bool,
) -> int:
    raw_dir = output_dir / ".raw_pdftoppm"
    if raw_dir.exists():
        shutil.rmtree(raw_dir)
    raw_dir.mkdir(parents=True)
    try:
        out_prefix = raw_dir / "page"
        command = [
            pdftoppm,
            "-png",
            "-scale-to-x",
            str(width),
            "-scale-to-y",
            "-1",
            str(pdf),
            str(out_prefix),
        ]
        run_command(command, "PDF 转 PNG", timeout=240)
        return move_numbered_pngs(raw_dir, output_dir, prefix, force)
    finally:
        shutil.rmtree(raw_dir, ignore_errors=True)


PDFKIT_JXA = r'''
ObjC.import('Foundation')
ObjC.import('AppKit')
ObjC.import('Quartz')

function run(argv) {
  const pdfPath = argv[0]
  const outDir = argv[1]
  const width = parseInt(argv[2] || '1920', 10)
  const prefix = argv[3] || 'page'
  const doc = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(pdfPath))
  if (!doc) throw new Error('Cannot open PDF: ' + pdfPath)

  $.NSFileManager.defaultManager.createDirectoryAtPathWithIntermediateDirectoriesAttributesError(outDir, true, $(), null)
  const count = Number(doc.pageCount)
  for (let i = 0; i < count; i++) {
    const page = doc.pageAtIndex(i)
    const bounds = page.boundsForBox($.kPDFDisplayBoxMediaBox)
    const height = Math.round(width * bounds.size.height / bounds.size.width)
    const image = page.thumbnailOfSizeForBox($.NSMakeSize(width, height), $.kPDFDisplayBoxMediaBox)
    const tiff = image.TIFFRepresentation
    const rep = $.NSBitmapImageRep.imageRepWithData(tiff)
    const png = rep.representationUsingTypeProperties($.NSBitmapImageFileTypePNG, $())
    const name = prefix + String(i + 1).padStart(3, '0') + '.png'
    png.writeToFileAtomically($(outDir).stringByAppendingPathComponent(name), true)
  }
  return String(count)
}
'''


def render_pdf_with_pdfkit(pdf: Path, output_dir: Path, width: int, prefix: str, force: bool) -> int:
    osascript = shutil.which("osascript")
    if not osascript:
        raise ConversionError("找不到 osascript，无法使用 macOS PDFKit 渲染。")

    raw_dir = output_dir / ".raw_pdfkit"
    if raw_dir.exists():
        shutil.rmtree(raw_dir)
    raw_dir.mkdir(parents=True)
    try:
        with tempfile.TemporaryDirectory(prefix="pptx-to-images-jxa.") as tmp:
            script = Path(tmp) / "pdf_to_png.jxa"
            script.write_text(PDFKIT_JXA, encoding="utf-8")
            run_command(
                [osascript, "-l", "JavaScript", str(script), str(pdf), str(raw_dir), str(width), "page"],
                "macOS PDFKit 转 PNG",
                timeout=240,
            )
        return move_numbered_pngs(raw_dir, output_dir, prefix, force)
    finally:
        shutil.rmtree(raw_dir, ignore_errors=True)


POWERPOINT_EXPORT_PS = r'''
param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputDir,
    [Parameter(Mandatory=$true)][int]$Width
)
$ErrorActionPreference = "Stop"
$powerPoint = $null
$presentation = $null

try {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($InputPath, $true, $false, $false)
    $ratio = [double]$presentation.PageSetup.SlideHeight / [double]$presentation.PageSetup.SlideWidth
    $height = [int][Math]::Round($Width * $ratio)
    $presentation.Export($OutputDir, "PNG", $Width, $height)
    [Console]::WriteLine($presentation.Slides.Count)
}
finally {
    if ($presentation -ne $null) {
        try { $presentation.Close() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($powerPoint -ne $null) {
        try { $powerPoint.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''


KEYNOTE_EXPORT_APPLESCRIPT = r'''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set outputAlias to POSIX file outputPath as alias
  tell application "Keynote"
    set doc to open POSIX file inputPath
    export doc to outputAlias as slide images with properties {image format:PNG}
    close doc saving no
  end tell
end run
'''


def export_with_powerpoint_windows(presentation: Path, output_dir: Path, width: int, prefix: str, force: bool) -> int:
    powershell = find_executable(["powershell.exe", "pwsh.exe", "powershell", "pwsh"])
    if not powershell:
        raise ConversionError("找不到 PowerShell。")

    raw_dir = output_dir / ".raw_powerpoint"
    if raw_dir.exists():
        shutil.rmtree(raw_dir)
    raw_dir.mkdir(parents=True)

    try:
        with tempfile.TemporaryDirectory(prefix="pptx-to-images-ps.") as tmp:
            script = Path(tmp) / "export_powerpoint.ps1"
            script.write_text(POWERPOINT_EXPORT_PS, encoding="utf-8")
            run_command(
                [
                    powershell,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-STA",
                    "-File",
                    str(script),
                    "-InputPath",
                    str(presentation),
                    "-OutputDir",
                    str(raw_dir),
                    "-Width",
                    str(width),
                ],
                "PowerPoint 导出 PNG",
                timeout=240,
            )
        return move_numbered_pngs(raw_dir, output_dir, prefix, force)
    finally:
        shutil.rmtree(raw_dir, ignore_errors=True)


def export_with_keynote_macos(presentation: Path, output_dir: Path, prefix: str, force: bool) -> int:
    if sys.platform != "darwin":
        raise ConversionError("Keynote 导出只支持 macOS。")
    if not Path("/Applications/Keynote.app").exists() and not (Path.home() / "Applications/Keynote.app").exists():
        raise ConversionError("找不到 Keynote。")

    osascript = shutil.which("osascript")
    if not osascript:
        raise ConversionError("找不到 osascript，无法自动控制 Keynote。")

    raw_dir = output_dir / ".raw_keynote"
    if raw_dir.exists():
        shutil.rmtree(raw_dir)
    raw_dir.mkdir(parents=True)

    try:
        with tempfile.TemporaryDirectory(prefix="pptx-to-images-keynote.") as tmp:
            script = Path(tmp) / "export_keynote.scpt"
            script.write_text(KEYNOTE_EXPORT_APPLESCRIPT, encoding="utf-8")
            run_command(
                [osascript, str(script), str(presentation), str(raw_dir)],
                "Keynote 导出 PNG",
                timeout=300,
            )
        return move_numbered_pngs(raw_dir, output_dir, prefix, force)
    finally:
        shutil.rmtree(raw_dir, ignore_errors=True)


def choose_engine(requested: str) -> EnginePlan:
    if requested == "powerpoint":
        return EnginePlan("powerpoint", "Windows PowerPoint COM")
    if requested == "libreoffice":
        return EnginePlan("libreoffice", "LibreOffice PDF + PNG renderer")
    if requested == "keynote":
        return EnginePlan("keynote", "macOS Keynote")
    if os.name == "nt":
        return EnginePlan("auto", "自动选择（Windows 优先 PowerPoint，失败后尝试 LibreOffice）")
    if sys.platform == "darwin":
        return EnginePlan("auto", "自动选择（LibreOffice，失败后尝试 Keynote）")
    return EnginePlan("auto", "自动选择（LibreOffice PDF + PNG renderer）")


def export_with_libreoffice(presentation: Path, output_dir: Path, width: int, prefix: str, force: bool) -> tuple[int, str]:
    soffice = find_soffice()
    if not soffice:
        raise ConversionError(
            "找不到 LibreOffice/soffice。\n"
            "请先安装 LibreOffice：\n"
            "  · Homebrew：brew install --cask libreoffice\n"
            "  · 或从官网下载：https://www.libreoffice.org/download/\n"
            "安装后重新运行即可。"
            "（Windows 装有 PowerPoint 时可改用 .cmd 启动器；macOS 装有 Keynote 时会自动兜底。）"
        )

    with tempfile.TemporaryDirectory(prefix="pptx-to-images.") as tmp:
        work_dir = Path(tmp)
        pdf = convert_to_pdf_with_libreoffice(presentation, work_dir, soffice)
        pdftoppm = find_pdftoppm()
        if pdftoppm:
            count = render_pdf_with_pdftoppm(pdf, output_dir, pdftoppm, width, prefix, force)
            return count, f"LibreOffice + pdftoppm ({Path(pdftoppm).name})"
        if sys.platform == "darwin":
            count = render_pdf_with_pdfkit(pdf, output_dir, width, prefix, force)
            return count, "LibreOffice + macOS PDFKit"
        raise ConversionError("已生成 PDF，但找不到 pdftoppm，无法把 PDF 每页渲染成 PNG。")


def export_presentation(
    presentation: Path,
    output_dir: Path,
    engine: EnginePlan,
    width: int,
    prefix: str,
    force: bool,
) -> tuple[int, str]:
    if engine.name == "auto":
        if os.name == "nt":
            try:
                count = export_with_powerpoint_windows(presentation, output_dir, width, prefix, force)
                return count, "Windows PowerPoint COM"
            except ConversionError as first_error:
                try:
                    return export_with_libreoffice(presentation, output_dir, width, prefix, force)
                except ConversionError as second_error:
                    raise ConversionError(
                        "PowerPoint 导出失败，LibreOffice 兜底也失败。\n\n"
                        f"PowerPoint：{first_error}\n\nLibreOffice：{second_error}"
                    ) from second_error
        if sys.platform == "darwin":
            try:
                return export_with_libreoffice(presentation, output_dir, width, prefix, force)
            except ConversionError as first_error:
                try:
                    count = export_with_keynote_macos(presentation, output_dir, prefix, force)
                    return count, "macOS Keynote"
                except ConversionError as second_error:
                    raise ConversionError(
                        "LibreOffice 导出失败，Keynote 兜底也失败。\n\n"
                        f"LibreOffice：{first_error}\n\nKeynote：{second_error}"
                    ) from second_error
        return export_with_libreoffice(presentation, output_dir, width, prefix, force)
    if engine.name == "powerpoint":
        if os.name != "nt":
            raise ConversionError("PowerPoint 自动导出当前只支持 Windows。macOS/Linux 请使用 LibreOffice。")
        count = export_with_powerpoint_windows(presentation, output_dir, width, prefix, force)
        return count, engine.detail
    if engine.name == "libreoffice":
        return export_with_libreoffice(presentation, output_dir, width, prefix, force)
    if engine.name == "keynote":
        count = export_with_keynote_macos(presentation, output_dir, prefix, force)
        return count, engine.detail
    raise ConversionError(f"未知导出引擎：{engine.name}")


def _open_path(path: Path) -> None:
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        elif os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="把 PPTX/PPT 每一页导出为 PNG 图片。无需第三方 Python 包，自动使用 PowerPoint、LibreOffice 或 Keynote。"
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help="PPTX/PPT 文件或包含演示文稿的文件夹。省略时进入拖入式引导。",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="输出文件夹。单个 PPTX 时直接输出到这里；多个 PPTX 时在其中创建子文件夹。",
    )
    parser.add_argument(
        "--engine",
        choices=("auto", "powerpoint", "libreoffice", "keynote"),
        default="auto",
        help="导出引擎。默认 auto：Windows 优先 PowerPoint；macOS 优先 LibreOffice，失败后尝试 Keynote。",
    )
    parser.add_argument(
        "--width",
        type=int,
        default=1920,
        help="输出图片目标宽度，默认 1920。高度按幻灯片比例自动计算。"
        "（Keynote 引擎按原生分辨率导出，此项不生效）",
    )
    parser.add_argument(
        "--prefix",
        default="page",
        help="输出图片名前缀，默认 page，生成 page001.png。",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="输入为文件夹时递归查找 PPTX/PPT。",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只预览将要转换的文件，不生成图片。",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="允许使用已有输出文件夹并覆盖同名图片。",
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="完成后打开输出文件夹。",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"pptx_to_images.py v{__version__}",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.width <= 0:
        print("--width 必须大于 0。", file=sys.stderr)
        return 2

    if os.name == "nt":
        for stream in (sys.stdout, sys.stderr, sys.stdin):
            try:
                stream.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
            except Exception:
                pass

    interactive = args.input is None
    if interactive:
        presentations = prompt_for_input(args.recursive)
        if presentations is None:
            print("已退出。")
            return 0
    else:
        try:
            presentations = resolve_input_presentations(args.input, args.recursive)
        except FileNotFoundError:
            print(f"输入路径不存在：{args.input}", file=sys.stderr)
            return 2
        except (OSError, ConversionError) as exc:
            print(str(exc), file=sys.stderr)
            return 2

    if not presentations:
        print("没有找到 PPTX/PPT 文件。", file=sys.stderr)
        return 1

    engine = choose_engine(args.engine)
    print("待转换文件：")
    for index, presentation in enumerate(presentations, start=1):
        print(f"  {index:03d}. {presentation}")
    print(f"\n导出引擎：{engine.detail}")
    print(f"目标宽度：{args.width}px")
    if engine.name == "keynote":
        print("注意：Keynote 引擎按幻灯片原生分辨率导出，--width 不生效。")

    if args.dry_run:
        print("\ndry-run 模式未生成图片。")
        return 0

    opened: list[Path] = []
    failures: list[Path] = []
    multiple = len(presentations) > 1
    total = len(presentations)
    for index, presentation in enumerate(presentations, start=1):
        try:
            output_dir = output_dir_for(presentation, args.output, multiple, args.force)
            print(f"\n[{index}/{total}] 正在导出：{presentation.name}")
            print(f"输出文件夹：{output_dir}")
            count, used_engine = export_presentation(
                presentation,
                output_dir,
                engine,
                args.width,
                args.prefix,
                args.force,
            )
        except ConversionError as exc:
            print(f"\n转换失败：{presentation}\n{exc}", file=sys.stderr)
            failures.append(presentation)
            continue

        print(f"完成：{count} 张 PNG（{used_engine}）")
        opened.append(output_dir)

    if failures:
        print(f"\n{len(failures)} 个文件转换失败：", file=sys.stderr)
        for presentation in failures:
            print(f"  - {presentation}", file=sys.stderr)

    if opened and (args.open or interactive):
        target = args.output.expanduser().resolve() if (args.output and multiple) else opened[0]
        _open_path(target)

    if interactive:
        try:
            input("\n按回车键退出……")
        except EOFError:
            pass

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
