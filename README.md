# Python-GUI-Mover
A Minimalist python GUI to move file because Windows's one is annoying.



# Ragilmalik’s Python GUI Mover 🗂️⚡

A sleek, minimal, **modern** Tkinter app to safely move files from one folder to another (top-level only — **no subfolders**), with smart metadata checks, automatic renaming, and gorgeous Excel logs.
Built for humans, not terminals. 😎

---

## ✨ Highlights

* **No surprises, ever**: Compares **filename**, **filetype**, **filesize**, and **filedate** (mtime). If **all match**, it **skips**. Otherwise it **moves** safely.
* **Zero overwrites**: If `file.jpg` exists with different metadata, the app renames to `file-1.jpg`, `file-2.jpg`, … until it finds a free name.
* **Simulation Only** by default: See exactly what would happen **before** you run for real.
* **Excel log (.xlsx)** with formatting:

  * Columns: **Timestamp**, **Action**, **Source Folder**, **Destination Folder**, **Filename**, **New Filename**, **File Creation Time**, **Size**, **Note**
  * **Timestamp format**: `DD/MM/YYYY HH:MM:SS`
  * **File Creation Time** format: `DD/MM/YYYY HH:MM:SS`
  * **Size** is rendered as text like `200KB`
  * **Bold** header row and **bold** summary row
  * 1 empty spacer row before the summary for neatness
* **Choose where logs go**: Save the log to **Source**, **Destination**, or a **Custom** folder.
* **Quality of life**:

  * **Open Last Saved Log File** (opens with default OS app)
  * **Clear Log Screen**
  * **Clear Log & Delete Last Log File**
* **Modern aesthetic**:

  * **Dark theme by default** (pure black base) and **Light theme** (pure white base)
  * Picker hover highlights switch to the **opposite** color (white on dark, black on light)
  * **Distinct button colors** for quick recognition
  * Subtle split-gradient header, clean typography, DPI awareness on Windows

---

## 🧠 How it decides: Move vs Skip (no subfolders)

1. Check if a same-named file already exists at the destination:

   * **No existing file** → **Move** it.
   * **Exists**:

     * If metadata (name, extension, size, mtime) **all match** → **Skip**.
     * Otherwise → **Move** and **auto-rename** to `name-1.ext`, `name-2.ext`, … (first free slot).
2. Only processes files in the **top level** of the source folder (no recursion).

---

## 📸 Screenshot

<img width="1365" height="732" alt="Screenshot_1" src="https://github.com/user-attachments/assets/f8e8a401-d21b-4e6c-b797-c78334912fe9" />


---

## 🚀 Getting Started

### Prerequisites

* **Python 3.8+**
* Windows/macOS/Linux
* One dependency: **openpyxl** (for Excel logs)

```bash
pip install openpyxl
```

### Run as a script

```bash
python gui.py
```

---

## 🧰 Build a single-file EXE (Windows)

> Requires **PyInstaller**: `pip install pyinstaller`

**CMD (one line):**

```cmd
pyinstaller --onefile --noconsole --clean --name "RagilmalikPythonGUIMover" gui.py
```

**PowerShell (one line):**

```powershell
pyinstaller --onefile --noconsole --clean --name "RagilmalikPythonGUIMover" gui.py
```

Optional: add an icon with `--icon youricon.ico`.

After a successful build you can **copy just the `.exe`** to another Windows machine and it will work (no Python required).
You may safely **delete** `build/`, `dist/` (after you grab the exe), and the `.spec` file if you don’t need custom PyInstaller tweaks.

---

## 📒 Log Format (Excel .xlsx)

Each run creates a timestamped workbook like `SmartFileMover-log-YYYY-MM-DD_HH-MM-SS.xlsx` with a single sheet:

**Columns**

* **Timestamp** — `DD/MM/YYYY HH:MM:SS`
* **Action** — `MOVED`, `MOVED_RENAMED`, `SKIP`, `DRYRUN_*`, `ERROR`, `INFO`, `SUMMARY`
* **Source Folder**
* **Destination Folder**
* **Filename**
* **New Filename** — empty if not renamed
* **File Creation Time** — `DD/MM/YYYY HH:MM:SS`
* **Size** — e.g., `200KB`
* **Note** — why it was moved/renamed/skipped

**Styling**

* Header row is **bold**
* One **empty row** at the end, then a **bold SUMMARY** row

---

## 🎛️ Controls & Options

* **Source Folder / Destination Folder** — pick the top-level folders
* **Log File Location (.xlsx)** — choose **Destination**, **Source**, or **Custom Folder**
* **Simulation Only** — on by default; uncheck to perform real moves
* **Run** — starts the job
* **Clear Log Screen** — clears the on-screen log area only
* **Clear Log & Delete Last Log File** — also deletes the last `.xlsx` produced this session
* **Open Last Saved Log File** — opens the most recent `.xlsx` with your OS default app
* **Theme** — Dark (pure black) / Light (pure white)

  * Picker hover highlights invert: white on dark, black on light
  * All text is the **opposite** of the theme base color

---

## 🧩 Tech Notes

* **GUI**: Tkinter + ttk
* **Logging**: openpyxl (`.xlsx`) with custom formats
* **Platform behaviors**:

  * **Open last log** uses `os.startfile` on Windows, `open` on macOS, `xdg-open` on Linux (with fallback to browser file URL)
  * File creation time uses `os.path.getctime`. On Linux this may reflect “metadata change time” rather than true creation time.

---

## ❓ FAQ

**Q: Why doesn’t it scan subfolders?**
A: Designed intentionally for speed and safety in top-level organization tasks. Keeping it flat prevents accidental deep moves. (If you want recursive mode later, see Roadmap 👇)

**Q: What happens if `file-1.jpg` already exists?**
A: The app keeps counting: `file-2.jpg`, `file-3.jpg`, … until it finds the first free name.

**Q: Can I trust “Simulation Only”?**
A: Yes. It performs all checks and tells you exactly what **would** happen — but doesn’t touch your files.

**Q: Where’s the log saved?**
A: Wherever you set (Source / Destination / Custom). You can open it instantly with the built-in button.

---

## 🗺️ Roadmap (ideas)

* Optional **recursive** mode with safe guards
* **Filters** (by extension/size/date)
* **Batch presets** (save & re-run common moves)
* Export **JSON**/CSV alongside Excel if needed
* **Drag & drop** folders onto the window

---

## 🧪 Development Notes

**Requirements file (optional)**

```txt
openpyxl>=3.1.0
```

**.gitignore (suggested)**

```
dist/
build/
*.spec
*.xlsx
__pycache__/
*.pyc
```

---

## 📄 License

MIT — do what you want, just don’t blame me if you yeet the wrong folder in Live mode. 😅
(Default run mode is **Simulation Only** for a reason!)

---

## 🙌 Credits

Crafted with care for clean file workflows.
If you find this useful, a ⭐ on GitHub makes my day!
