# PPTX 转图片器 / PPTX to Images

[English Documentation](#english-documentation) | [Documentation française](#documentation-française)

## 中文说明

### PPTX 转图片器

`pptx_to_images` 用于把 PowerPoint `.pptx` / `.ppt` 演示文稿逐页导出为 PNG 图片。输出文件按页码命名为 `page001.png`、`page002.png`，可直接交给参考项目 `images_to_pptx` 再合成为 PPTX。

本工具面向“拆页成图片 → 修改图片 → 再合成 PPTX”的工作流：先把每页幻灯片渲染成图片，后续可以批量修图、标注、压缩，最后用 `images_to_pptx` 按编号顺序重新生成演示文稿。

### 功能

| 类型 | 支持内容 |
|------|----------|
| 输入文件 | `.pptx`、`.ppt` |
| 输出图片 | PNG，默认宽度 `1920px`，高度按幻灯片比例自动计算 |
| 批量处理 | 可传入文件夹，自动处理其中的 PPTX/PPT 文件 |
| 输出命名 | `page001.png`、`page002.png`、`page003.png` |
| 渲染引擎 | Windows PowerPoint、LibreOffice、macOS Keynote 兜底 |
| 免配置 | Python 入口不需要第三方包；启动器提供拖拽式引导 |

### 运行入口

本项目提供三个运行入口：

- Python 源码入口：`pptx_to_images.py`
- macOS 启动器：`pptx_to_images.command`
- Windows 启动器：`pptx_to_images.cmd`

### 运行环境

| 入口 | 运行要求 |
|------|----------|
| `pptx_to_images.py` | Python 3.8+；不需要第三方 Python 包 |
| `pptx_to_images.command` | macOS + Python 3；同目录保留 `pptx_to_images.py` |
| `pptx_to_images.cmd` | Windows + PowerShell；优先使用 Microsoft PowerPoint 导出 |

实际渲染 PPTX 需要系统中已有一个演示文稿渲染引擎：

| 系统 | 推荐环境 |
|------|----------|
| Windows | Microsoft PowerPoint；Python 源码入口也可使用 LibreOffice |
| macOS | LibreOffice；无 `pdftoppm` 时自动使用 macOS PDFKit；LibreOffice 不可用时尝试 Keynote |
| Linux | LibreOffice + `pdftoppm` |

本项目不需要安装 Python 第三方包，也不需要虚拟环境。若已安装参考项目可用的 Office / LibreOffice 环境，本工具可以直接搭配使用。

### 安装 LibreOffice

macOS / Linux 默认使用 LibreOffice 渲染。如果尚未安装，导出时会报「找不到 LibreOffice/soffice」，请先安装：

**macOS**

```bash
# 方式一：Homebrew
brew install --cask libreoffice

# 方式二：从官网下载安装包
# https://www.libreoffice.org/download/
```

首次启动 LibreOffice 可能较慢（构建字体缓存等），属正常现象；导出过程会显示进度与耗时，不会卡在空白屏。

**Linux**

```bash
# Debian / Ubuntu
sudo apt install libreoffice poppler-utils

# Fedora
sudo dnf install libreoffice poppler-utils
```

其中 `poppler-utils` 提供 `pdftoppm`，用于把 PDF 每页渲染成 PNG。

**Windows**

装有 Microsoft PowerPoint 时直接用 `pptx_to_images.cmd` 即可，无需 LibreOffice；也可安装 LibreOffice 配合 Python 入口使用。

**没有 LibreOffice 时的兜底**：macOS 上如果装了 Keynote，工具会自动改用 Keynote 导出。注意 Keynote 按幻灯片原生分辨率导出（`--width` 不生效），且首次运行需在「系统设置 > 隐私与安全性 > 自动化」里授权。

### macOS 使用方式

`pptx_to_images.command` 是 macOS 双击入口，会调用同目录下的 `pptx_to_images.py`。

1. 双击 `pptx_to_images.command`
2. 如果系统提示“无法验证开发者”，右键点击文件，选择“打开”，再确认运行
3. 将 `.pptx` 文件拖入窗口，按回车
4. 图片会生成在 PPTX 文件旁边的 `文件名_images/` 目录

也可以在终端中运行：

```bash
./pptx_to_images.command presentation.pptx
python3 pptx_to_images.py presentation.pptx --open
```

### Windows 使用方式

`pptx_to_images.cmd` 是 Windows 单文件启动入口。双击运行时会提示拖入或输入 `.pptx` / `.ppt` 文件路径；也可以将 PPTX 文件直接拖到 `.cmd` 文件上运行。

```cmd
pptx_to_images.cmd presentation.pptx
```

Windows 启动器使用 PowerShell 自动控制 PowerPoint 导出 PNG，因此适合装有 Microsoft PowerPoint 的电脑。若要使用 LibreOffice 兜底，请使用 Python 源码入口。

### Python 源码使用方式

```bash
# 导出单个 PPTX
python3 pptx_to_images.py presentation.pptx

# 指定输出目录
python3 pptx_to_images.py presentation.pptx -o presentation_images

# 指定图片宽度
python3 pptx_to_images.py presentation.pptx --width 2560

# 处理文件夹中的 PPTX/PPT
python3 pptx_to_images.py ./slides --recursive

# 预览将要转换的文件，不生成图片
python3 pptx_to_images.py ./slides --dry-run
```

可用选项：

| 选项 | 说明 |
|------|------|
| 位置参数 | PPTX/PPT 文件或文件夹；省略时进入拖入式引导 |
| `-o` / `--output` | 指定输出文件夹 |
| `--engine auto\|powerpoint\|libreoffice\|keynote` | 指定导出引擎，默认自动 |
| `--width` | 输出 PNG 目标宽度，默认 `1920` |
| `--prefix` | 输出文件名前缀，默认 `page` |
| `--recursive` | 输入为文件夹时递归查找 |
| `--dry-run` | 只预览，不生成图片 |
| `--force` | 使用已有输出文件夹并覆盖同名图片 |
| `--open` | 完成后打开输出文件夹 |

### 输出目录结构

默认情况下，输出目录会创建在 PPTX 文件旁边，目录名为 `<文件名>_images/`。如果输出目录已存在，会自动追加 `_2`、`_3` 等后缀，除非显式使用 `--force`。

```text
presentation_images/
├── page001.png
├── page002.png
└── page003.png
```

处理多个 PPTX 文件并指定同一个输出目录时：

```text
all_images/
├── presentation1_images/
│   ├── page001.png
│   └── page002.png
└── presentation2_images/
    ├── page001.png
    └── page002.png
```

### 工作原理

PowerPoint `.pptx` 文件本身是 ZIP/XML 包，但“逐页导出为图片”需要真实渲染引擎。本工具的流程如下：

1. 解析输入路径，确定要处理的 PPTX/PPT 文件
2. 自动选择可用渲染引擎
3. Windows 优先调用 PowerPoint COM 直接导出 PNG
4. macOS/Linux 优先调用 LibreOffice 将 PPTX 转为 PDF
5. 使用 `pdftoppm` 或 macOS PDFKit 将 PDF 每页渲染成 PNG
6. 按页码整理输出文件名为 `page001.png`、`page002.png`

### 与 `images_to_pptx` 搭配

本工具负责：

```text
PPTX -> page001.png / page002.png / ...
```

参考项目 `images_to_pptx` 负责：

```text
page001.png / page002.png / ... -> PPTX
```

两个工具可以组合成完整的拆分与重建流程。

### 发布记录

详见 [docs/](docs/) 目录下各版本发布记录。

### 开源协议

MIT

---

## English Documentation

### PPTX to Images

`pptx_to_images` exports every slide in a PowerPoint `.pptx` / `.ppt` file to PNG images. Output files are named `page001.png`, `page002.png`, and so on, so they can be fed directly into the companion `images_to_pptx` project.

### Features

| Type | Supported Content |
|------|-------------------|
| Input files | `.pptx`, `.ppt` |
| Output images | PNG, default width `1920px`, height calculated from slide aspect ratio |
| Batch mode | Accepts a folder and processes PPTX/PPT files inside it |
| Naming | `page001.png`, `page002.png`, `page003.png` |
| Engines | Windows PowerPoint, LibreOffice, macOS Keynote fallback |
| Setup | No third-party Python packages or virtual environment required |

### Runtime Entry Points

- Python source entry point: `pptx_to_images.py`
- macOS launcher: `pptx_to_images.command`
- Windows launcher: `pptx_to_images.cmd`

### Requirements

| Entry point | Requirement |
|-------------|-------------|
| `pptx_to_images.py` | Python 3.8+; no third-party Python packages |
| `pptx_to_images.command` | macOS + Python 3; keep `pptx_to_images.py` in the same folder |
| `pptx_to_images.cmd` | Windows + PowerShell; uses Microsoft PowerPoint when available |

Rendering requires an installed presentation engine: PowerPoint on Windows, LibreOffice on macOS/Linux, or Keynote as a macOS fallback.

### Installing LibreOffice

On macOS / Linux the default renderer is LibreOffice. If it is not installed, export fails with “找不到 LibreOffice/soffice” (LibreOffice/soffice not found). Install it first:

**macOS**

```bash
# Option 1: Homebrew
brew install --cask libreoffice

# Option 2: download the installer
# https://www.libreoffice.org/download/
```

LibreOffice's first launch can be slow (building font caches, etc.) — this is normal. The export shows progress and elapsed time instead of a blank screen.

**Linux**

```bash
# Debian / Ubuntu
sudo apt install libreoffice poppler-utils

# Fedora
sudo dnf install libreoffice poppler-utils
```

`poppler-utils` provides `pdftoppm`, which renders each PDF page to PNG.

**Windows**

With Microsoft PowerPoint installed, just use `pptx_to_images.cmd` — LibreOffice is not required. You can also install LibreOffice and use the Python entry point.

**Fallback without LibreOffice**: on macOS, if Keynote is installed the tool automatically exports with Keynote (Keynote exports at the slide's native resolution, so `--width` has no effect, and the first run needs Automation permission under System Settings → Privacy & Security → Automation).

### Usage

```bash
python3 pptx_to_images.py presentation.pptx
python3 pptx_to_images.py presentation.pptx -o presentation_images
python3 pptx_to_images.py presentation.pptx --width 2560
python3 pptx_to_images.py ./slides --recursive
```

On macOS:

```bash
./pptx_to_images.command presentation.pptx
```

On Windows:

```cmd
pptx_to_images.cmd presentation.pptx
```

### Output Structure

```text
presentation_images/
├── page001.png
├── page002.png
└── page003.png
```

### How It Works

Slide-to-image export needs a rendering engine. The tool uses PowerPoint COM on Windows when possible. On macOS/Linux it converts the presentation to PDF with LibreOffice, then renders each PDF page to PNG with `pdftoppm` or macOS PDFKit.

### Release Notes

See per-version release notes in the [docs/](docs/) directory.

### License

MIT

---

## Documentation française

### PPTX vers images

`pptx_to_images` exporte chaque diapositive d’un fichier PowerPoint `.pptx` / `.ppt` en images PNG. Les fichiers générés sont nommés `page001.png`, `page002.png`, etc., afin de pouvoir être réutilisés directement avec le projet compagnon `images_to_pptx`.

### Fonctionnalités

| Type | Contenu pris en charge |
|------|------------------------|
| Fichiers d’entrée | `.pptx`, `.ppt` |
| Images de sortie | PNG, largeur par défaut `1920px`, hauteur calculée selon le ratio de la diapositive |
| Traitement par lot | Accepte un dossier contenant des fichiers PPTX/PPT |
| Nommage | `page001.png`, `page002.png`, `page003.png` |
| Moteurs | PowerPoint sous Windows, LibreOffice, Keynote en secours sur macOS |
| Installation | Aucun paquet Python tiers ni environnement virtuel nécessaire |

### Points d’entrée

- Script Python source : `pptx_to_images.py`
- Lanceur macOS : `pptx_to_images.command`
- Lanceur Windows : `pptx_to_images.cmd`

### Prérequis

| Point d’entrée | Prérequis |
|----------------|-----------|
| `pptx_to_images.py` | Python 3.8+ ; aucun paquet Python tiers |
| `pptx_to_images.command` | macOS + Python 3 ; conserver `pptx_to_images.py` dans le même dossier |
| `pptx_to_images.cmd` | Windows + PowerShell ; utilise Microsoft PowerPoint si disponible |

Le rendu nécessite un moteur de présentation installé : PowerPoint sous Windows, LibreOffice sous macOS/Linux, ou Keynote comme solution de secours sur macOS.

### Installer LibreOffice

Sous macOS / Linux, le moteur de rendu par défaut est LibreOffice. S'il n'est pas installé, l'export échoue avec « 找不到 LibreOffice/soffice ». Installez-le d'abord :

**macOS**

```bash
# Option 1 : Homebrew
brew install --cask libreoffice

# Option 2 : télécharger l'installateur
# https://www.libreoffice.org/download/
```

Le premier lancement de LibreOffice peut être lent (construction des caches de polices, etc.), c'est normal. L'export affiche la progression et le temps écoulé.

**Linux**

```bash
# Debian / Ubuntu
sudo apt install libreoffice poppler-utils

# Fedora
sudo dnf install libreoffice poppler-utils
```

`poppler-utils` fournit `pdftoppm`, qui rend chaque page PDF en PNG.

**Windows**

Avec Microsoft PowerPoint installé, utilisez simplement `pptx_to_images.cmd` — LibreOffice n'est pas nécessaire. Vous pouvez aussi installer LibreOffice et utiliser le point d'entrée Python.

**Solution de secours sans LibreOffice** : sous macOS, si Keynote est installé, l'outil exporte automatiquement avec Keynote (Keynote exporte à la résolution native de la diapositive, donc `--width` est sans effet, et la première exécution nécessite l'autorisation d'automatisation dans Réglages Système → Confidentialité et sécurité → Automatisation).

### Utilisation

```bash
python3 pptx_to_images.py presentation.pptx
python3 pptx_to_images.py presentation.pptx -o presentation_images
python3 pptx_to_images.py presentation.pptx --width 2560
python3 pptx_to_images.py ./slides --recursive
```

Sur macOS :

```bash
./pptx_to_images.command presentation.pptx
```

Sur Windows :

```cmd
pptx_to_images.cmd presentation.pptx
```

### Structure de sortie

```text
presentation_images/
├── page001.png
├── page002.png
└── page003.png
```

### Fonctionnement

L’export de diapositives en images nécessite un moteur de rendu. L’outil utilise PowerPoint COM sous Windows lorsque c’est possible. Sous macOS/Linux, il convertit la présentation en PDF avec LibreOffice, puis rend chaque page PDF en PNG avec `pdftoppm` ou PDFKit sous macOS.

### Notes de publication

Consultez les notes de version dans le dossier [docs/](docs/).

### Licence

MIT
