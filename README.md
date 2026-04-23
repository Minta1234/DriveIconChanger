# DriveIconChanger

# 💾 Drive & Folder Icon Setter

> **Windows utility to set custom icons for drives and folders — with a built-in crop editor, live Explorer refresh, and cross-platform icon support.**

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-blue?logo=windows)
![Version](https://img.shields.io/badge/Version-13.0-brightgreen)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Admin](https://img.shields.io/badge/Requires-Administrator-red)

---

## ✨ Features

### 💾 Drive Icon Tool
- Set custom icons for **any drive** — C:\\, D:\\, USB drives, external HDDs, RAM disks
- Writes to **Windows Registry** (`HKLM\SOFTWARE\...\DriveIcons`) for instant effect
- Supports **system drives** (C:\\) with ProgramData fallback — no root write errors
- Creates `autorun.inf` for icon display on other Windows PCs
- Cross-platform icon files included for portable drives:
  - 🐧 **Linux KDE/Dolphin** — `.directory` + `drive_icon.png`
  - 🐧 **Linux GNOME/Nemo/Thunar** — `.xdg-volume-info`
  - 🍎 **macOS** — `.VolumeIcon.icns` (structure prepared)
- **Auto Eject** after applying (USB/Removable drives)
- Removes icons cleanly — no leftover blank icon bug

### 📁 Folder Icon Tool
- Set custom icons for **any folder** on your PC
- Creates `desktop.ini` with `[.ShellClassInfo]` automatically
- Sets required folder `+S` (System) attributes
- Hides icon files with `+H+S` attributes (optional)
- Stores icons in `%ProgramData%\FolderIcons` for safe keeping
- Clean removal with full `desktop.ini` cleanup

### 🖼️ Built-in Crop / Edit Editor
- **Drag to pan**, **scroll to zoom** on the source image
- Real-time **256px preview** + small sizes (48px, 32px, 16px)
- Background options: **Transparent**, **White**, **Black**, **Circle crop**
- High-quality **LANCZOS** resampling for final output
- Exports multi-resolution `.ico` (256 / 128 / 64 / 48 / 32 / 16 px)

### ⚡ Smart Shell Refresh
- Uses `SHChangeNotify` API — **no Explorer kill/restart** for drive icons
- Sends `DRIVEREMOVED` → `DRIVEADD` sequence to force full icon re-read
- Flushes icon cache via `ie4uinit.exe -ClearIconCache`
- Fixes the common **blank icon bug** caused by stale `desktop.ini` entries

---

---

## 🚀 Getting Started

### Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 or Windows 11 |
| Python | 3.8 or newer |
| Pillow | Auto-installed on first run |
| Privileges | **Administrator** (for Registry writes) |

### Installation

```bash
# Clone the repository
git clone https://github.com/Minta1234/DriveIconChanger.git
cd DriveIconChanger

# (Optional) Install Pillow manually
pip install Pillow
```

### Running

**Right-click → Run as Administrator** is recommended for full functionality.

```bash
python DriveIconSetter-V2.py
```

Or double-click `DriveIconSetter-V2.py` → choose "Run as administrator" if prompted.

---

## 🛠️ How It Works

### Setting a Drive Icon

```
Your image (PNG/JPG/ICO/etc.)
        │
        ▼
  [Crop Editor]  ← drag, zoom, background options
        │
        ▼
  drive_icon.ico  (multi-res: 256/128/64/48/32/16px)
        │
        ├──→ Registry: HKLM\SOFTWARE\...\DriveIcons\{Letter}\DefaultIcon
        ├──→ %ProgramData%\DriveIcons\drive_{Letter}.ico
        ├──→ {Drive}\.icons\drive_icon.ico
        ├──→ {Drive}\.icons\drive_icon.png       (Linux KDE)
        ├──→ {Drive}\.icons\drive_icon_512.png   (Linux HiDPI)
        ├──→ {Drive}\desktop.ini                 (Windows Explorer)
        ├──→ {Drive}\autorun.inf                 (Other PCs)
        ├──→ {Drive}\.directory                  (Linux KDE/Dolphin)
        └──→ {Drive}\.xdg-volume-info            (Linux GNOME/Nemo)
```

### Setting a Folder Icon

```
Your image
        │
        ▼
  [Crop Editor]
        │
        ▼
  folder_icon.ico
        │
        ├──→ %ProgramData%\FolderIcons\{FolderName}\folder_icon.ico
        ├──→ {Folder}\desktop.ini  → [.ShellClassInfo] IconResource=...
        └──→ attrib +S {Folder}    (System flag — required by Windows)
```

---

## 📁 File Structure

```
drive-folder-icon-setter/
│
├── DriveIconSetter-V2.py          # Main application (single-file)
└── README.md
```

> The app is intentionally single-file for easy distribution.

---

## ⚙️ Technical Details

| Item | Detail |
|---|---|
| GUI framework | `tkinter` (built into Python) |
| Image processing | `Pillow (PIL)` — auto-installed |
| Registry access | `winreg` (Python stdlib) |
| Shell notification | `ctypes` → `SHChangeNotify` |
| ICO export | Multi-res: 256/128/64/48/32/16 px |
| Theme | Catppuccin Mocha |
| Threading | Background pipeline, UI stays responsive |

---

## 🐛 Known Limitations

- **Windows only** — the app will exit on Linux/macOS (by design; use the Linux edition for cross-platform management)
- System drive (C:\\) icon changes require a **PC restart** to fully take effect in some cases
- GNOME Nautilus (modern) requires `gio set` to be run from Linux — the Windows edition writes `.xdg-volume-info` as a best-effort fallback
- `.VolumeIcon.icns` (macOS) structure is written but full ICNS generation is not yet implemented

---

## 🗺️ Roadmap

- [ ] Full `.icns` generation for macOS icons
- [ ] Batch apply icon to multiple folders at once
- [ ] Preset icon library / icon pack support
- [ ] Linux edition (GNOME `gio set` support)
- [ ] CLI mode (headless / scripting support)

---

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you'd like to change.

1. Fork the repository
2. Create your branch: `git checkout -b feature/my-feature`
3. Commit your changes: `git commit -m 'Add my feature'`
4. Push: `git push origin feature/my-feature`
5. Open a Pull Request

---

## 📄 License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) for details.

---

## 👤 Author

**Minta**  
GitHub: [@Minta1234](https://github.com/Minta1234)  
Website: [minta0077.online](https://minta0077.online)

---

<p align="center">Made with ❤️ and way too many Registry keys</p>
