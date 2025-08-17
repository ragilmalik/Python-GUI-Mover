# Python-GUI-Mover
A Minimalist python GUI to move file because Windows's one is annoying.



# Ragilmalik’s Python GUI Mover 🗂️⚡

A sleek, minimal, **modern** Tkinter app to safely move files from one folder to another (top-level only — **no subfolders**), with smart metadata checks, automatic renaming, and gorgeous Excel logs.  
Built for humans, not terminals. 😎

---

## ✨ Highlights

- **No surprises, ever**: Compares **filename**, **filetype**, **filesize**, and **filedate** (mtime). If **all match**, it **skips**. Otherwise it **moves** safely.
- **Zero overwrites**: If `file.jpg` exists with different metadata, the app renames to `file-1.jpg`, `file-2.jpg`, … until it finds a free name.
- **Simulation Only** by default: See exactly what would happen **before** you run for real.
- **Excel log (.xlsx)** with formatting:
  - Columns: **Timestamp**, **Action**, **Source Folder**, **Destination Folder**, **Filename**, **New Filename**, **File Creation Time**, **Size**, **Note**
  - **Timestamp format**: `DD/MM/YYYY HH:MM:SS`
  - **File Creation Time** format: `DD/MM/YYYY HH:MM:SS`
  - **Size** is rendered as text like `200KB`
  - **Bold** header row and **bold** summary row
  - 1 empty spacer row before the summary for neatness
- **Choose where logs go**: Save the log to **Source**, **Destination**, or a **Custom** folder.
- **Quality of life**:
  - **Open Last Saved Log File** (opens with default OS app)
  - **Clear Log Screen**
  - **Clear Log & Delete Last Log File**
- **Modern aesthetic**:
  - **Dark theme by default** (pure black base) and **Light theme** (pure white base)
  - Picker hover highlights switch to the **opposite** color (white on dark, black on light)
  - **Distinct button colors** for quick recognition
  - Subtle split-gradient header, clean typography, DPI awareness on Windows

---

## 🧠 How it decides: Move vs Skip (no subfolders)

1. Check if a same-named file already exists at the destination:
   - **No existing file** → **Move** it.
   - **Exists**:
     - If metadata (name, extension, size, mtime) **all match** → **Skip**.
     - Otherwise → **Move** and **auto-rename** to `name-1.ext`, `name-2.ext`, … (first free slot).
2. Only processes files in the **top level** of the source folder (no recursion).

---

## 📸 Screenshot



---

## 🚀 Getting Started

### Prerequisites
- **Python 3.8+**
- Windows/macOS/Linux
- One dependency: **openpyxl** (for Excel logs)

```bash
pip install openpyxl
