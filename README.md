
````markdown
# Ragilmalik’s Python GUI Mover 🗂️⚡

A sleek, minimal, **modern** Tkinter app to safely move files from one folder to another (top-level or **recursive**), with **content-accurate duplicate detection (SHA-256)**, per-folder auto-renaming, long-path support on Windows, and gorgeous Excel logs.  
Built for humans, not terminals. 😎

---

## ✨ Highlights

- **True duplicate detection** (tiered & fast):  
  1) If the destination name doesn’t exist → **Move**.  
  2) If it exists and **sizes differ** → **Rename & move** (definitely different).  
  3) If it exists and **sizes match** → compare **SHA-256** of both files.  
     - **Hashes equal** → **Skip** (identical content).  
     - **Hashes differ** → **Rename & move**.
- **Recursive (preserve structure)**: Optional checkbox. Keeps the directory tree relative to the source.
- **Skip hidden files/folders** (optional): Default is **off** (hidden are moved).
- **File type filter**: Include-only (e.g. `jpg,png,mp4`). Leave empty = all files.
- **Zero overwrites**: If `file.jpg` collides and differs, renames to `file-1.jpg`, `file-2.jpg`, … per **containing folder**.
- **Long path support (Windows)**: Uses `\\?\` automatically under the hood for reliability.
- **Simulation Only** by default: See exactly what would happen **before** you run for real.
- **Excel log (.xlsx)** with formatting and clarity:
  - Columns: **Timestamp**, **Action**, **Source Folder**, **Destination Folder**, **Filename**, **New Filename**, **File Creation Time**, **Size**, **Note**
  - **Timestamp** & **File Creation Time** format: `DD/MM/YYYY HH:MM:SS`
  - **Size** rendered as text like `200KB`
  - **Bold** header row + **bold** summary row, with a spacer row for neatness
  - **Always shows full destination path + filename** in the **Filename** column, formatted as `{Fullpath}/{Filename}` (forward slashes on all OS)
  - **Destination Folder** column holds the **final destination directory** (absolute)
- **Quality of life**:
  - **Open Last Saved Log File** (opens with default OS app)
  - **Clear Log Screen**
  - **Clear Log & Delete Last Log File**
  - **Pause / Resume / Stop** (safe)
  - **Generate tollback script (Undo)** — builds a `.bat` (Windows) or `.sh` (macOS/Linux) from the last **Live** run

---

## 🧠 How it decides: Move vs Skip

1. **No existing dest file** → **Move**.  
2. **Existing dest file**:
   - **Different size** → **Rename & move** (`name-1.ext`, `name-2.ext`, …).
   - **Same size** → compute **SHA-256** for both:
     - **Same hash** → **Skip** (identical content).
     - **Different hash** → **Rename & move**.
3. **Recursive mode** preserves **relative paths** under the destination.
4. Renames are **per destination directory** (not global).

---

## 📸 Screenshot

> <img width="1365" height="729" alt="Screenshot_1" src="https://github.com/user-attachments/assets/d1a17a78-a505-4fcb-a566-3670e65951d8" />


---

## 🚀 Getting Started

### Prerequisites
- **Python 3.8+**
- Windows/macOS/Linux
- Dependency: **openpyxl** (for Excel logs)

```bash
pip install openpyxl
````

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
* **Source Folder** — absolute path
* **Destination Folder** — **absolute final directory** of the file
* **Filename** — **absolute destination path including filename**, formatted `{Fullpath}/{Filename}`
* **New Filename** — empty if not renamed
* **File Creation Time** — `DD/MM/YYYY HH:MM:SS`
* **Size** — e.g., `200KB`
* **Note** — e.g., “Different size; renamed”, “Identical content (SHA-256)”, etc.

**Styling**

* Header row is **bold**
* One **empty row** at the end, then a **bold SUMMARY** row

---

## 🎛️ Controls & Options

* **Source Folder / Destination Folder**
* **Log File Location (.xlsx)** — **Destination**, **Source**, or **Custom Folder**
* **Simulation Only** — default ON (safe)
* **Recursive (preserve structure)** — default OFF
* **Skip hidden files/folders** — default OFF
* **File type filter** — include-only list like `jpg,png,mp4`
* **Run**, **Pause**, **Resume**, **Stop**
* **Clear Log Screen**
* **Clear Log & Delete Last Log File**
* **Open Last Saved Log File**
* **Generate tollback script (Undo)**
* **Theme** — Dark (pure black) / Light (pure white)

  * Picker hover highlights invert: white on dark, black on light
  * All text is the **opposite** of the theme base color

---

## 🧩 Tech Notes

* **GUI**: Tkinter + ttk
* **Logging**: openpyxl (`.xlsx`) with custom formats
* **Duplicate detection**: **SHA-256** hashing with 1MB chunks, used **only** when sizes match and names collide (with a **hash cache** to avoid recomputing).
* **Windows**:

  * Long paths handled via `\\?\` internally.
  * On-screen/Excel paths normalized to forward slashes for readability.
* **Preserves timestamps**:

  * Same-volume moves are renames (preserve times).
  * Cross-volume moves fall back to copy2 + delete, preserving **mtime/atime**.

---

## ❓ FAQ

**Does the log always show full paths?**
Yes — both the on-screen log and the Excel **Filename** column show the **absolute destination path + filename** in `{Fullpath}/{Filename}` format.

**Is hashing slow?**
Hashing reads the whole file, so it’s used **only** when a destination name already exists **and sizes match**. The built-in **hash cache** minimizes re-hashing repeats.

**What if `file-1.jpg` already exists?**
The app keeps counting: `file-2.jpg`, `file-3.jpg`, … until it finds the first free name — **per destination directory**.

**Can I trust “Simulation Only”?**
Yes. It performs all checks (including SHA-256 where applicable) and reports exactly what **would** happen — without touching your files.

---

## 🗺️ Roadmap (ideas)

* Volume/depth caps & free-space preflight UI
* Advanced exclusions and safe flatten mode
* Batch presets
* Export JSON manifest alongside Excel
* Drag & drop folders onto the window

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
