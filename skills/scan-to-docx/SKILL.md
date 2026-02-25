---
name: scan-to-docx
description: 將掃描圖片轉換成 .docx，使用 macOS Vision OCR（繁體中文優先）。用法：/scan-to-docx <來源圖片或資料夾> <輸出路徑>。輸出路徑若為 .docx 則合併成單一檔，若為資料夾則每張圖各輸出一個 .docx。
disable-model-invocation: true
allowed-tools: Bash
---

# scan-to-docx

將掃描圖片（PNG、JPG、TIFF 等）透過 macOS Vision OCR 辨識（繁體中文優先），並輸出為 Word .docx 文件。

## 用法

```
$ARGUMENTS
```

請解析 `$ARGUMENTS`，格式為：
```
<來源> <輸出>
```

- `<來源>`：單張圖片路徑，或包含圖片的資料夾路徑
- `<輸出>`：
  - 若副檔名為 `.docx` → **合併模式**：所有圖片合併成一個 .docx
  - 若為資料夾路徑 → **獨立模式**：每張圖各輸出一個 .docx

## 執行步驟

1. 確認來源路徑存在
2. 確認輸出路徑的父目錄存在（若不存在則建立）
3. 確認 Python 依賴已安裝（首次使用時自動安裝）：

```bash
pip3 install --quiet Pillow python-docx pyobjc-framework-Vision pyobjc-framework-Quartz
```

4. 執行腳本：

```bash
python3 ~/.claude/skills/scan-to-docx/scripts/scan_to_docx.py <來源> <輸出>
```

5. 回報每張圖的處理結果（OCR 統計、圖片數量）

## 範例

合併模式（多張圖 → 一個檔）：
```
/scan-to-docx ~/Desktop/掃描圖片/ ~/Desktop/輸出.docx
```

獨立模式（每張圖各一個檔）：
```
/scan-to-docx ~/Desktop/掃描圖片/ ~/Desktop/輸出資料夾/
```

單張圖片：
```
/scan-to-docx ~/Desktop/page1.png ~/Desktop/page1.docx
```

## 注意事項

- 僅支援 macOS（需要 Vision framework 與 Quartz）
- 支援格式：`.png` `.jpg` `.jpeg` `.webp` `.tiff` `.tif` `.bmp`
- OCR 語言優先順序：繁體中文 → 簡體中文 → 英文
- 年代頁（大字年份）自動整頁輸出為圖片
- 自動偵測跨欄與同欄照片區域，並配對圖說
