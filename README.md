# scan-to-docx

A [Claude Code](https://claude.ai/code) skill that converts scanned images to Word `.docx` files using macOS Vision OCR, with Traditional Chinese as the primary language.

[繁體中文說明](#繁體中文) | [English](#english)

---

## 繁體中文

### 功能特色

- OCR 辨識優先順序：繁體中文 → 簡體中文 → 英文
- 自動偵測跨欄與同欄照片區域
- 自動配對圖說與最近的照片
- 左欄文字先、右欄文字後輸出
- 年代頁（大字年份）自動整頁保留為圖片
- 兩種輸出模式：**合併**（所有圖片 → 一個 `.docx`）或**獨立**（每張圖各一個 `.docx`）

### 系統需求

- **僅支援 macOS**（需要 Vision framework 與 Quartz）
- Python 3
- pip 依賴（首次使用時自動安裝）：
  - `Pillow`
  - `python-docx`
  - `pyobjc-framework-Vision`
  - `pyobjc-framework-Quartz`

### 安裝

將 skill 目錄複製到 Claude Code 個人 skills 資料夾：

```bash
git clone https://github.com/javaing/scan-to-docx.git
cp -r scan-to-docx/skills/scan-to-docx ~/.claude/skills/scan-to-docx
```

### 使用方式

```
/scan-to-docx <來源> <輸出>
```

| 來源 | 輸出 | 模式 |
|------|------|------|
| 單張圖片 | `output.docx` | 合併模式 |
| 圖片資料夾 | `output.docx` | 全部合併成一個檔 |
| 圖片資料夾 | `output-folder/` | 每張圖各輸出一個 `.docx` |

**範例：**

```bash
# 單張圖片
/scan-to-docx ~/Desktop/page1.png ~/Desktop/page1.docx

# 多張圖片合併成一個檔
/scan-to-docx ~/Desktop/掃描圖片/ ~/Desktop/輸出.docx

# 每張圖各輸出一個 .docx
/scan-to-docx ~/Desktop/掃描圖片/ ~/Desktop/輸出資料夾/
```

### 支援格式

`.png` `.jpg` `.jpeg` `.webp` `.tiff` `.tif` `.bmp`

---

## English

### Features

- OCR priority: Traditional Chinese → Simplified Chinese → English
- Auto-detects photo regions (cross-column and single-column)
- Matches captions to their nearest photos
- Outputs left column before right column
- Year-title pages are preserved as full-page images
- Two output modes: **merge** (all images → one `.docx`) or **individual** (one `.docx` per image)

### Requirements

- **macOS only** (requires Vision framework and Quartz)
- Python 3
- pip dependencies (auto-installed on first use):
  - `Pillow`
  - `python-docx`
  - `pyobjc-framework-Vision`
  - `pyobjc-framework-Quartz`

### Installation

```bash
git clone https://github.com/javaing/scan-to-docx.git
cp -r scan-to-docx/skills/scan-to-docx ~/.claude/skills/scan-to-docx
```

### Usage

```
/scan-to-docx <source> <output>
```

| Source | Output | Mode |
|--------|--------|------|
| Single image | `output.docx` | Merge |
| Folder of images | `output.docx` | Merge all into one file |
| Folder of images | `output-folder/` | Individual `.docx` per image |

**Examples:**

```bash
# Single image
/scan-to-docx ~/Desktop/page1.png ~/Desktop/page1.docx

# Merge multiple images into one file
/scan-to-docx ~/Desktop/scans/ ~/Desktop/output.docx

# One .docx per image
/scan-to-docx ~/Desktop/scans/ ~/Desktop/output-folder/
```

### Supported Formats

`.png` `.jpg` `.jpeg` `.webp` `.tiff` `.tif` `.bmp`

---

## License

MIT
