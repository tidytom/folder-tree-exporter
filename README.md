# Folder Tree Exporter

A small Windows tool that lists every folder and subfolder inside a folder you pick (Dropbox, OneDrive or a local drive) and saves the list as an Excel or CSV file, for anyone who needs an overview of a large folder structure.

> Screenshot: add `docs/screenshot.png` (placeholder line — leave it exactly like this for GUI apps)

## What it does
- Scans a folder and all of its subfolders, at any depth.
- Saves one row per folder, with the parent folders in the columns `Level1`, `Level2`, and so on.
- Saves as Excel (`.xlsx`, with a header row, filters and frozen top row) or as CSV.
- Shows a progress window with the number of folders found and the current path, with a **Cancel** button.
- Reads folder names only, so it works with online-only Dropbox and OneDrive folders without downloading files.
- Writes a log of folders it had to skip, and why.

## What you need
- Windows 10 or 11 (it also runs on macOS and Linux).
- Python 3.8 or newer, with tkinter (included in the python.org installer for Windows).
- `openpyxl` for Excel output (in `requirements.txt`). Without it, the tool saves CSV files.

## Install
1. Download this project (green **Code** button, then **Download ZIP**) and unzip it.
2. Open a Command Prompt in the project folder.
3. Run:
   ```
   pip install -r requirements.txt
   ```

## How to use
1. Start the tool:
   ```
   python folder_tree_exporter.py
   ```
2. Click **OK** on the "Pick Folder" message, then choose the folder to scan.
3. Choose where to save the list and its name (the default is `<folder name>_folders.xlsx`). Pick **Excel Workbook** or **CSV (Comma delimited)**.
4. On Windows, the **Options** window asks about **Use Windows long-path prefix (\\?\)**. Leave it ticked and click **Start Scan**.
5. Watch the **Scanning...** window. Click **Cancel** to stop early; the folders found so far are still saved.
6. When it finishes, the **Done** message shows where the file and the log were saved, how many folders were found and how many were skipped.

## Settings and options
- **Use Windows long-path prefix (\\?\)**: lets the tool read folders whose full path is longer than 260 characters. On by default.
- **File type**: `.xlsx` needs openpyxl. If you choose `.xlsx` without openpyxl, the list is saved as `.csv` with the same name instead.

Example of the result:

```
Level1   | Level2     | Level3
---------|------------|--------
Dropbox  |            |
Dropbox  | Projects   |
Dropbox  | Projects   | Alpha
Dropbox  | Projects   | Beta
```

## Safety notes
- The tool only reads folder names. It does not change, move or delete anything in the scanned folder, and it does not open or download files.
- It writes two files: the list you named, and a log next to it named `<list name>.scan_log.txt`.
- It never replaces a file without asking: the save window asks before it replaces the file you chose, and the log (or a CSV saved instead of an Excel file) gets a new name such as `name (1).csv` if the name is already taken.
- Nothing is sent anywhere.

## Troubleshooting
- **"Excel output not available"**: openpyxl is missing. Run `pip install -r requirements.txt` in the project folder, or continue and save a CSV file.
- **Some folders are listed as skipped in the log**: they may be online-only folders the sync program has not made available yet, folders moved during the scan, or folders you have no permission to read. Run the scan again after the sync program has finished.
- **Folders with very long paths are missing**: make sure **Use Windows long-path prefix** is ticked.
- **"No module named tkinter"**: install Python from python.org and keep the "tcl/tk and IDLE" option ticked.

## How it works
The tool walks the folder tree with `os.scandir`, checking only whether each entry is a folder (`is_dir` without following links), so cloud placeholders are not downloaded. The scan runs on a worker thread and sends progress to the tkinter window through a queue, so **Cancel** always responds. Paths can use the Windows `\\?\` extended-length prefix to pass the 260-character limit. The rows are written with `openpyxl` (Excel) or the `csv` module.

## License

MIT. See [LICENSE](LICENSE).
