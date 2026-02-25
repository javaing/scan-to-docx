# scan-to-docx

A [Claude Code](https://claude.ai/code) skill that converts scanned images to Word `.docx` files using macOS Vision OCR, with Traditional Chinese as the primary language.

## Features

- OCR priority: Traditional Chinese → Simplified Chinese → English
- Auto-detects photo regions (cross-column and single-column)
- Matches captions to their nearest photos
- Outputs left column before right column
- Year-title pages are preserved as full-page images
- Two output modes: **merge** (all images → one `.docx`) or **individual** (one `.docx` per image)

## Requirements

- **macOS only** (requires Vision framework and Quartz)
- Python 3
- pip dependencies (auto-installed on first use):
  - `Pillow`
  - `python-docx`
  - `pyobjc-framework-Vision`
  - `pyobjc-framework-Quartz`

## Installation

Copy the skill directory to your personal Claude Code skills folder:

```bash
cp -r skills/scan-to-docx ~/.claude/skills/scan-to-docx
```

Or clone and symlink:

```bash
git clone https://github.com/javaing/scan-to-docx.git
ln -s "$(pwd)/scan-to-docx/skills/scan-to-docx" ~/.claude/skills/scan-to-docx
```

## Usage

```
/scan-to-docx <source> <output>
```

| Source | Output | Mode |
|--------|--------|------|
| Single image | `output.docx` | Merge |
| Folder of images | `output.docx` | Merge all into one file |
| Folder of images | `output-folder/` | Individual `.docx` per image |

### Examples

```bash
# Single image
/scan-to-docx ~/Desktop/page1.png ~/Desktop/page1.docx

# Merge multiple images into one file
/scan-to-docx ~/Desktop/scans/ ~/Desktop/output.docx

# One .docx per image
/scan-to-docx ~/Desktop/scans/ ~/Desktop/output-folder/
```

## Supported Formats

`.png` `.jpg` `.jpeg` `.webp` `.tiff` `.tif` `.bmp`

## License

MIT
