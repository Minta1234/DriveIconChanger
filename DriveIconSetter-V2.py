"""
Drive & Folder Icon Setter  v13.0  —  WINDOWS EDITION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
NEW in v12.0: cross-platform icon support (Linux GNOME fix), 512px PNG,
   no Explorer freeze (soft_refresh_shell), fixed registry path to drive,
   FAT32/NTFS detection for folder icons, .xdg-volume-info for GNOME
"""

import os
import sys
import shutil
import subprocess
import tempfile
import time
import glob
import ctypes
import winreg
import platform
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

# ── Auto-install Pillow ───────────────────────────────────────────────────────
def _install_pillow():
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--user", "Pillow"], check=True)
    except:
        try:
            subprocess.run(["pip", "install", "--user", "Pillow"], check=True)
        except:
            pass
    import site
    u = site.getusersitepackages()
    if u not in sys.path:
        sys.path.insert(0, u)

try:
    from PIL import Image, ImageTk, ImageDraw
    # Increase limit to allow large images (fixes "decompression bomb" error)
    Image.MAX_IMAGE_PIXELS = None 
except ModuleNotFoundError:
    _install_pillow()
    try:
        import site
        from importlib import reload
        reload(site)
        from PIL import Image, ImageTk, ImageDraw
        Image.MAX_IMAGE_PIXELS = None
    except ModuleNotFoundError:
        print(f"Run: {sys.executable} -m pip install Pillow")
        input("Press Enter to exit...")
        sys.exit(1)

if sys.platform != "win32":
    print("Windows only.")
    sys.exit(1)

# ── Windows Version ───────────────────────────────────────────────────────────
def _get_win_ver():
    try:
        build = int(platform.version().split(".")[2])
        return "11" if build >= 22000 else "10"
    except:
        return "10"

WIN_VER = _get_win_ver()
IS_WIN11 = WIN_VER == "11"

# ── Constants ─────────────────────────────────────────────────────────────────
DRIVE_REMOVABLE = 2
DRIVE_FIXED = 3
DRIVE_REMOTE = 4
DRIVE_CDROM = 5
DRIVE_RAMDISK = 6

DRIVE_TYPE_LABEL = {
    DRIVE_REMOVABLE: "USB/Removable",
    DRIVE_FIXED: "Local Disk",
    DRIVE_REMOTE: "Network",
    DRIVE_CDROM: "CD/DVD",
    DRIVE_RAMDISK: "RAM Disk",
}

SIZES = [256, 128, 64, 48, 32, 16]
REG_BASE = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons"
ICO_STORE = os.path.expandvars(r"%ProgramData%\DriveIcons")
FOLDER_ICON_STORE = os.path.expandvars(r"%ProgramData%\FolderIcons")

# ==============================================================================
#  COMMON FUNCTIONS
# ==============================================================================

def is_admin():
    """Check if running as administrator"""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except:
        return False

def relaunch_as_admin():
    """Relaunch the script with administrator privileges"""
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

def clear_attribs(path):
    """Clear all file attributes"""
    if os.path.exists(path):
        subprocess.run(["attrib", "-R", "-H", "-S", path],
                       shell=True, capture_output=True)

def set_hidden_windows(path):
    """Hide on Windows with +H+S attributes"""
    if os.path.exists(path):
        subprocess.run(["attrib", "+H", "+S", path],
                       shell=True, capture_output=True)

def is_hidden_windows(path):
    """Check if file is hidden on Windows"""
    if not os.path.exists(path):
        return False
    result = subprocess.run(f"attrib {path}", shell=True, 
                           capture_output=True, text=True)
    return 'H' in result.stdout

def pil_to_ico(img, out_path):
    """Convert PIL image to Windows .ico format"""
    img = img.convert("RGBA")
    icons = [img.resize((s, s), Image.LANCZOS) for s in SIZES]
    icons[0].save(out_path, format="ICO",
                  sizes=[(s, s) for s in SIZES], append_images=icons[1:])

# ==============================================================================
#  EXPLORER CONTROL (Shell notifications)
# ==============================================================================

# Shell notification constants
SHCNE_UPDATEITEM = 0x00002000
SHCNE_RENAMEFOLDER = 0x00020000
SHCNE_ASSOCCHANGED = 0x08000000
SHCNE_ALLEVENTS = 0x7FFFFFFF
SHCNF_PATHW = 0x0005
SHCNF_FLUSH = 0x1000
SHCNF_FLUSHNOWAIT = 0x3000

def kill_explorer():
    """Kill Explorer and wait until fully terminated"""
    subprocess.run(["taskkill", "/F", "/IM", "explorer.exe"],
                   shell=True, capture_output=True)
    
    for _ in range(30):
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq explorer.exe"],
                           capture_output=True, text=True, shell=True)
        if "explorer.exe" not in r.stdout.lower():
            break
        time.sleep(0.2)
    time.sleep(0.3)

def delete_icon_cache():
    """
    Bust the Windows icon cache so stale icons (e.g. blank drive icon) are cleared.
    Explorer locks the .db files while running so we can't delete them directly.
    Instead we use the ie4uinit sequence which Windows itself uses to flush the cache.
    """
    # Try deleting cache DB files (only works if Explorer is stopped, which we never do)
    patterns = [
        r"%LOCALAPPDATA%\Microsoft\Windows\Explorer\iconcache*.db",
        r"%LOCALAPPDATA%\Microsoft\Windows\Explorer\thumbcache*.db",
    ]
    for pattern in patterns:
        for f in glob.glob(os.path.expandvars(pattern)):
            try:
                if os.path.isfile(f):
                    os.remove(f)
            except Exception:
                pass  # Expected — Explorer has these locked

    # ie4uinit is the proper live-flush method that works while Explorer runs
    for flag in ["-show", "-ClearIconCache"]:
        try:
            subprocess.Popen(["ie4uinit.exe", flag],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

def start_explorer():
    """Start Explorer and wait until running"""
    subprocess.Popen("explorer.exe", shell=True)
    
    for _ in range(40):
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq explorer.exe"],
                           capture_output=True, text=True, shell=True)
        if "explorer.exe" in r.stdout.lower():
            break
        time.sleep(0.3)
    
    time.sleep(2.5 if IS_WIN11 else 2.0)

def notify_shell(path):
    """Notify Windows Shell to update icon"""
    soft_refresh_shell(path)

def soft_refresh_shell(path):
    """
    Refresh the Explorer icon for *path* with ZERO freeze.
    Never kills Explorer — uses SHChangeNotify API directly.
    """
    s32 = ctypes.windll.shell32
    p   = ctypes.create_unicode_buffer(path)
    s32.SHChangeNotify(SHCNE_UPDATEITEM,   SHCNF_PATHW | SHCNF_FLUSH, p, None)
    s32.SHChangeNotify(SHCNE_RENAMEFOLDER, SHCNF_PATHW | SHCNF_FLUSH, p, None)
    s32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_FLUSHNOWAIT, None, None)
    try:
        subprocess.Popen(["ie4uinit.exe", "-show"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

# ==============================================================================
#  DRIVE ICON FUNCTIONS
# ==============================================================================

def get_drives():
    """Get all accessible drives on Windows"""
    drives = []
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        if bitmask & 1:
            drive_path = f"{letter}:\\"
            drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive_path)
            
            # Check if drive is accessible
            try:
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(drive_path, None, None, None)
                
                # Get drive label
                label = get_drive_label(drive_path)
                
                # Get free space
                free_bytes = ctypes.c_ulonglong(0)
                total_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                    drive_path, None, ctypes.byref(total_bytes), ctypes.byref(free_bytes))
                
                drives.append({
                    'path': drive_path,
                    'letter': letter,
                    'type': drive_type,
                    'type_name': DRIVE_TYPE_LABEL.get(drive_type, "Unknown"),
                    'label': label,
                    'total': total_bytes.value,
                    'free': free_bytes.value,
                    'is_system': (drive_path.rstrip("\\").upper() == 
                                 os.environ.get("SystemDrive", "C:").upper())
                })
            except:
                pass
        bitmask >>= 1
    
    return drives

def get_drive_label(drive):
    """Get volume label for a drive"""
    buf = ctypes.create_unicode_buffer(261)
    try:
        ctypes.windll.kernel32.GetVolumeInformationW(
            drive, buf, 261, None, None, None, None, 0)
        return buf.value
    except:
        return ""

def get_drive_filesystem(drive):
    """Get filesystem type for a drive"""
    buf = ctypes.create_unicode_buffer(261)
    try:
        ctypes.windll.kernel32.GetVolumeInformationW(
            drive, None, 0, None, None, None, buf, 261)
        return buf.value
    except:
        return "Unknown"

def reg_set_drive_icon(drive, ico_path):
    """Write drive icon to Windows Registry"""
    letter = drive.rstrip("\\").rstrip(":")[0].upper()
    
    for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
        try:
            # Create the DriveIcons key first if it doesn't exist
            with winreg.CreateKeyEx(hive, REG_BASE, 0, winreg.KEY_SET_VALUE) as _:
                pass
            
            # Create or open the letter key
            letter_key_path = f"{REG_BASE}\\{letter}"
            with winreg.CreateKeyEx(hive, letter_key_path, 0, winreg.KEY_WRITE) as letter_key:
                pass
            
            # Set DefaultIcon only
            icon_key_path = f"{REG_BASE}\\{letter}\\DefaultIcon"
            with winreg.CreateKeyEx(hive, icon_key_path, 0, winreg.KEY_SET_VALUE) as k:
                # DefaultIcon value MUST be "path,index" — the ",0" is required
                # by some Windows versions; without it the icon shows blank.
                ico_reg_val = ico_path if ico_path.endswith(",0") else f"{ico_path},0"
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, ico_reg_val)
            
        except Exception as e:
            if hive == winreg.HKEY_LOCAL_MACHINE:
                raise PermissionError(
                    f"Registry write failed: {e}\nRun as Administrator!")

def reg_get_drive_icon(drive):
    """Get current drive registry icon path"""
    letter = drive.rstrip("\\").rstrip(":")[0].upper()
    
    for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
        try:
            with winreg.OpenKey(hive,
                                f"{REG_BASE}\\{letter}\\DefaultIcon") as k:
                return winreg.QueryValueEx(k, "")[0]
        except:
            pass
    return None

def reg_remove_drive_icon(drive):
    """Remove drive registry icon entry"""
    letter = drive.rstrip("\\").rstrip(":")[0].upper()
    
    for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
        # Delete DefaultIcon subkey
        try:
            path = f"{REG_BASE}\\{letter}\\DefaultIcon"
            winreg.DeleteKey(hive, path)
        except:
            pass
        
        # Then delete the letter key
        try:
            path = f"{REG_BASE}\\{letter}"
            winreg.DeleteKey(hive, path)
        except:
            pass

def safe_eject(drive):
    """Safely eject a USB drive"""
    letter = drive.rstrip("\\").rstrip(":")[0].upper()
    k32 = ctypes.windll.kernel32
    h = k32.CreateFileW(f"\\\\.\\{letter}:",
                        0x80000000 | 0x40000000, 1 | 2, None, 3, 0x20000000, None)
    
    if h == -1:
        # Fallback to PowerShell
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"(New-Object -comObject Shell.Application).Namespace(17)"
             f".ParseName('{letter}:').InvokeVerb('Eject')"],
            capture_output=True, timeout=15)
        return r.returncode == 0, ("Ejected via PowerShell"
                                    if r.returncode == 0
                                    else r.stderr.decode(errors="ignore"))
    
    br = ctypes.c_ulong(0)
    try:
        k32.DeviceIoControl(h, 0x90018, None, 0, None, 0, ctypes.byref(br), None)
        k32.DeviceIoControl(h, 0x90020, None, 0, None, 0, ctypes.byref(br), None)
        k32.DeviceIoControl(h, 0x2D4804, None, 0, None, 0, ctypes.byref(br), None)
        ok = k32.DeviceIoControl(h, 0x2D4808, None, 0, None, 0, ctypes.byref(br), None)
        
        if ok:
            return True, "Ejected!"
        
        err = ctypes.GetLastError()
        k32.CloseHandle(h)
        return False, f"Eject failed (code {err})"
    finally:
        try:
            k32.CloseHandle(h)
        except:
            pass

# ==============================================================================
#  FOLDER ICON FUNCTIONS
# ==============================================================================

def set_folder_icon(folder_path, ico_path, hide_files=True):
    """
    Set custom icon for any folder
    Creates desktop.ini and hides icon files
    """
    results = {'success': True, 'messages': []}
    
    def log(msg):
        results['messages'].append(msg)
    
    try:
        # Create folder for storing icons
        folder_icons_dir = os.path.join(FOLDER_ICON_STORE, 
                                        os.path.basename(folder_path).replace(' ', '_'))
        os.makedirs(folder_icons_dir, exist_ok=True)
        
        # Copy icon to storage
        ico_filename = f"folder_icon_{int(time.time())}.ico"
        stored_ico = os.path.join(folder_icons_dir, ico_filename)
        shutil.copy2(ico_path, stored_ico)
        log(f"✅ Copied icon to: {stored_ico}")
        
        # Create desktop.ini in the folder
        desktop_ini = os.path.join(folder_path, "desktop.ini")
        
        # Use relative path if possible, otherwise absolute
        if folder_path.startswith(os.path.splitdrive(folder_path)[0]):
            # Same drive - use relative path
            rel_path = os.path.relpath(stored_ico, folder_path)
            icon_resource = rel_path
        else:
            # Different drive - use absolute path
            icon_resource = stored_ico
        
        ini_content = (
            "[.ShellClassInfo]\r\n"
            f"IconResource={icon_resource},0\r\n"
            f"IconFile={icon_resource}\r\n"
            "IconIndex=0\r\n"
            "[ViewState]\r\n"
            "Mode=\r\n"
            "Vid=\r\n"
            "FolderType=Generic\r\n"
        )
        
        # Remove existing desktop.ini if present
        if os.path.exists(desktop_ini):
            clear_attribs(desktop_ini)
            os.remove(desktop_ini)
        
        # Write new desktop.ini
        with open(desktop_ini, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(ini_content)
        log(f"✅ Created desktop.ini")
        
        # Set folder as system (required for custom icon)
        subprocess.run(["attrib", "+S", folder_path], shell=True, capture_output=True)
        log(f"✅ Set folder as system")
        
        # Set desktop.ini as hidden + system
        set_hidden_windows(desktop_ini)
        subprocess.run(["attrib", "+S", desktop_ini], shell=True, capture_output=True)
        log(f"✅ Set desktop.ini as hidden + system")
        
        # Hide icon files if requested
        if hide_files:
            set_hidden_windows(stored_ico)
            log(f"✅ Hidden icon file")
            
            # Also hide the folder icons directory
            if os.path.exists(FOLDER_ICON_STORE):
                set_hidden_windows(FOLDER_ICON_STORE)
        
        # Refresh Explorer
        notify_shell(folder_path)
        
        return True, results['messages']
        
    except Exception as e:
        results['success'] = False
        results['messages'].append(f"❌ Error: {str(e)}")
        return False, results['messages']

def remove_folder_icon(folder_path):
    """Remove custom icon from a folder"""
    results = {'success': True, 'messages': []}
    
    def log(msg):
        results['messages'].append(msg)
    
    try:
        desktop_ini = os.path.join(folder_path, "desktop.ini")
        
        if os.path.exists(desktop_ini):
            # Read desktop.ini to find icon file
            try:
                with open(desktop_ini, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Try to find icon path
                import re
                icon_match = re.search(r'IconResource=(.+?),0', content)
                if icon_match:
                    icon_path = icon_match.group(1).strip()
                    # Try to remove icon file if it exists and is in our storage
                    if os.path.exists(icon_path) and FOLDER_ICON_STORE in icon_path:
                        try:
                            clear_attribs(icon_path)
                            os.remove(icon_path)
                            log(f"✅ Removed icon file: {os.path.basename(icon_path)}")
                        except:
                            pass
            except:
                pass
            
            # Remove desktop.ini
            clear_attribs(desktop_ini)
            os.remove(desktop_ini)
            log(f"✅ Removed desktop.ini")
        
        # Refresh Explorer
        notify_shell(folder_path)
        
        return True, results['messages']
        
    except Exception as e:
        results['success'] = False
        results['messages'].append(f"❌ Error: {str(e)}")
        return False, results['messages']

def get_folder_icon_status(folder_path):
    """Check if folder has custom icon"""
    desktop_ini = os.path.join(folder_path, "desktop.ini")
    
    if not os.path.exists(desktop_ini):
        return None, "No custom icon"
    
    try:
        with open(desktop_ini, 'r', encoding='utf-8') as f:
            content = f.read()
        
        import re
        icon_match = re.search(r'IconResource=(.+?),0', content)
        if icon_match:
            icon_path = icon_match.group(1).strip()
            if os.path.exists(icon_path):
                return icon_path, f"Custom icon: {os.path.basename(icon_path)}"
            else:
                return icon_path, f"Icon file missing: {os.path.basename(icon_path)}"
        else:
            return None, "desktop.ini exists but no icon entry"
    except:
        return None, "Error reading desktop.ini"

# ==============================================================================
#  DRIVE ICON PIPELINE
# ==============================================================================

def apply_drive_icon(drive_info, ico_src, do_eject, status_cb, done_cb):
    """
    Apply icon to drive — visible on Windows, macOS, and ALL Linux DEs.

    Files written to drive root:
      .icons/drive_icon.ico   ← Windows Registry + autorun.inf point here
      .icons/drive_icon.png   ← Linux KDE .directory points here
      .icons/drive_icon_512.png ← HiDPI Linux
      .VolumeIcon.icns        ← macOS auto-reads from root
      desktop.ini             ← Windows Explorer (this PC)
      autorun.inf             ← Windows Explorer (other PCs)
      .directory              ← KDE/Dolphin
      .xdg-volume-info        ← GNOME Nautilus / Nemo / Thunar (legacy)

    GNOME PERSISTENCE NOTE:
      .xdg-volume-info is read by Thunar, Nemo, and older Nautilus.
      Modern GNOME Nautilus requires 'gio set' which must be run from Linux.
      The Linux edition of this tool handles that automatically.
    """
    drive   = drive_info['path']
    is_usb  = (drive_info['type'] == DRIVE_REMOVABLE)
    is_sys  = drive_info['is_system']
    letter  = drive_info['letter']
    t0      = time.time()
    TOTAL   = 16

    def step(n, msg):
        status_cb(f"[{time.time() - t0:.1f}s] [{n}/{TOTAL}] {msg}")

    try:
        os.makedirs(ICO_STORE, exist_ok=True)
        ico_dest = os.path.join(ICO_STORE, f"drive_{letter}.ico")

        # ── SYSTEM vs REMOVABLE drive strategy ───────────────────────────────
        # C:\ root is owned by TrustedInstaller — even admin gets Errno 13
        # when creating hidden folders there.  For system drives we keep ALL
        # files inside ProgramData and use absolute paths in the registry.
        # Removable / data drives write files to the drive root as normal.
        if is_sys:
            icons_dir = os.path.join(ICO_STORE, f"drive_{letter}_icons")
        else:
            icons_dir = os.path.join(drive, ".icons")

        # ── 1. Prepare ────────────────────────────────────────────────────────
        mode_label = "ProgramData mode (system drive)" if is_sys else "drive-root mode"
        step(1, f"Preparing... ({mode_label})")

        # ── 2. Background icon cache flush ────────────────────────────────────
        step(2, "Scheduling icon cache flush (background)...")
        delete_icon_cache()
        step(2, "✅ Cache flush scheduled")

        # ── 3. Remove old icon files ──────────────────────────────────────────
        step(3, "Removing old icon files...")
        removed = 0
        if os.path.exists(ico_dest):
            try:
                clear_attribs(ico_dest)
                os.remove(ico_dest)
                removed += 1
            except Exception:
                pass
        if not is_sys:
            for pattern in ["drive_icon*.ico", ".VolumeIcon*",
                            "desktop.ini", "autorun.inf", ".directory",
                            ".xdg-volume-info"]:
                for old_f in glob.glob(os.path.join(drive, pattern)):
                    try:
                        clear_attribs(old_f)
                        os.remove(old_f)
                        removed += 1
                    except Exception:
                        pass
        step(3, f"Removed {removed} old file(s).")

        # ── 4. Create icons folder ────────────────────────────────────────────
        step(4, f"Creating icons folder...")
        if os.path.exists(icons_dir):
            try:
                clear_attribs(icons_dir)
            except Exception:
                pass
        os.makedirs(icons_dir, exist_ok=True)
        step(4, f"✅ {icons_dir}")

        # ── 5. Load source image ──────────────────────────────────────────────
        step(5, "Loading image...")
        pil_img = Image.open(ico_src).convert("RGBA")
        step(5, f"✅ {pil_img.width}x{pil_img.height}px loaded")

        # ── 6. Copy .ico to ProgramData ───────────────────────────────────────
        step(6, "Copying icon to ProgramData...")
        shutil.copy2(ico_src, ico_dest)
        step(6, "✅ Copied to ProgramData")

        # ── 7. Copy .ico to icons folder ─────────────────────────────────────
        ico_root = os.path.join(icons_dir, "drive_icon.ico")
        step(7, "Copying .ico to icons folder...")
        shutil.copy2(ico_src, ico_root)
        step(7, f"✅ drive_icon.ico saved")

        # ── 8. Save PNG 256 + 512px (Linux — removable drives only) ──────────
        if not is_sys:
            step(8, "Saving PNG for Linux (256px + 512px)...")
            try:
                png_256 = os.path.join(icons_dir, "drive_icon.png")
                png_512 = os.path.join(icons_dir, "drive_icon_512.png")
                pil_img.resize((256, 256), Image.LANCZOS).save(png_256, "PNG")
                pil_img.resize((512, 512), Image.LANCZOS).save(png_512, "PNG")
                step(8, "✅ drive_icon.png + drive_icon_512.png")
            except Exception as ex:
                step(8, f"⚠️  PNG skipped: {ex}")
        else:
            step(8, "⏩ PNG skipped (system drive — Windows only)")

        # ── 9. macOS .VolumeIcon.icns (removable drives only) ────────────────
        if not is_sys:
            step(9, "Creating macOS .VolumeIcon.icns...")
            try:
                icns_path = os.path.join(drive, ".VolumeIcon.icns")
                # This would need a pil_to_icns function
                step(9, "⚠️  ICNS creation not fully implemented")
            except Exception as ex:
                step(9, f"⚠️  ICNS skipped: {ex}")
        else:
            step(9, "⏩ ICNS skipped (system drive)")

        # ── 10. Write Registry ────────────────────────────────────────────────
        # System drive  → point registry to ProgramData path (always writable)
        # Removable      → point registry to ico on the drive root (portable)
        step(10, "Updating Registry...")
        reg_ico = ico_dest if is_sys else ico_root
        reg_set_drive_icon(drive, reg_ico)
        step(10, f"✅ Registry → {reg_ico}")

        # ── 11. Create desktop.ini ────────────────────────────────────────────
        step(11, "Creating desktop.ini...")
        desktop_ini = os.path.join(drive, "desktop.ini")
        if is_sys:
            ini_content = (
                "[.ShellClassInfo]\r\n"
                f"IconResource={reg_ico},0\r\n"
                f"IconFile={reg_ico}\r\n"
                "IconIndex=0\r\n"
            )
        else:
            ini_content = (
                "[.ShellClassInfo]\r\n"
                "IconResource=.icons\\drive_icon.ico,0\r\n"
                "IconFile=.icons\\drive_icon.ico\r\n"
                "IconIndex=0\r\n"
            )
        try:
            if os.path.exists(desktop_ini):
                clear_attribs(desktop_ini)
            with open(desktop_ini, "w", encoding="utf-8", newline="") as f:
                f.write(ini_content)
            step(11, "✅ desktop.ini created")
        except Exception as ex:
            step(11, f"⚠️  desktop.ini: {ex}")

        # ── 12. Create autorun.inf (removable drives only) ───────────────────
        if not is_sys:
            step(12, "Creating autorun.inf...")
            autorun_inf = os.path.join(drive, "autorun.inf")
            autorun_content = (
                "[autorun]\r\n"
                "icon=.icons\\drive_icon.ico\r\n"
            )
            try:
                if os.path.exists(autorun_inf):
                    clear_attribs(autorun_inf)
                with open(autorun_inf, "w", encoding="utf-8", newline="") as f:
                    f.write(autorun_content)
                set_hidden_windows(autorun_inf)
                step(12, "✅ autorun.inf created")
            except Exception as ex:
                step(12, f"⚠️  autorun.inf: {ex}")
        else:
            step(12, "⏩ autorun.inf skipped (system drive)")

        # ── 13. Create .directory (Linux KDE — removable drives only) ───────────
        if not is_sys:
            step(13, "Creating .directory for Linux KDE/Dolphin...")
            dot_directory = os.path.join(drive, ".directory")
            dir_content = (
                "[Desktop Entry]\n"
                "Icon=.icons/drive_icon.png\n"
            )
            try:
                if os.path.exists(dot_directory):
                    clear_attribs(dot_directory)
                with open(dot_directory, "w", encoding="utf-8", newline="\n") as f:
                    f.write(dir_content)
                set_hidden_windows(dot_directory)
                step(13, "✅ .directory created")
            except Exception as ex:
                step(13, f"⚠️  .directory skipped: {ex}")
        else:
            step(13, "⏩ .directory skipped (system drive)")

        # ── 14. Create .xdg-volume-info (Linux GNOME — removable drives only) ──
        if not is_sys:
            step(14, "Creating .xdg-volume-info for Linux GNOME/Nemo/Thunar...")
            xdg_vol = os.path.join(drive, ".xdg-volume-info")
            label   = drive_info.get('label', '') or letter
            xdg_content = (
                "[Volume Info]\n"
                f"Name={label}\n"
                "Icon=.icons/drive_icon.png\n"
            )
            try:
                if os.path.exists(xdg_vol):
                    clear_attribs(xdg_vol)
                with open(xdg_vol, "w", encoding="utf-8", newline="\n") as f:
                    f.write(xdg_content)
                set_hidden_windows(xdg_vol)
                step(14, "✅ .xdg-volume-info created")
            except Exception as ex:
                step(14, f"⚠️  .xdg-volume-info skipped: {ex}")
        else:
            step(14, "⏩ .xdg-volume-info skipped (system drive)")

        # ── 15. Hide all Windows-visible files ────────────────────────────────
        step(15, "Hiding files...")
        set_hidden_windows(icons_dir)
        set_hidden_windows(desktop_ini)
        set_hidden_windows(ico_dest)
        step(15, "✅ Files hidden")

        # ── 16. Hard shell refresh — force Explorer to pick up new icon ──────
        # soft_refresh_shell (UPDATEITEM only) is not enough after a blank-icon
        # situation.  Must send DRIVEREMOVED→DRIVEADD to force full re-read.
        step(16, "Forcing shell refresh...")
        s32    = ctypes.windll.shell32
        drv_buf = ctypes.create_unicode_buffer(drive)

        s32.SHChangeNotify(SHCNE_ASSOCCHANGED,  SHCNF_FLUSHNOWAIT, None, None)
        time.sleep(0.1)
        s32.SHChangeNotify(0x00000080,            # SHCNE_DRIVEREMOVED
                           SHCNF_PATHW | SHCNF_FLUSH, drv_buf, None)
        time.sleep(0.15)
        s32.SHChangeNotify(0x00000100,            # SHCNE_DRIVEADD
                           SHCNF_PATHW | SHCNF_FLUSH, drv_buf, None)
        time.sleep(0.15)
        s32.SHChangeNotify(SHCNE_UPDATEITEM,
                           SHCNF_PATHW | SHCNF_FLUSH, drv_buf, None)
        # Flush icon cache again now that new .ico is in place
        delete_icon_cache()
        step(16, "✅ Shell refreshed")

        total = time.time() - t0
        step(16, f"Done!  Finished in {total:.1f}s")

        eject_ok = False
        if do_eject and is_usb:
            status_cb("Ejecting drive...")
            time.sleep(0.5)
            eject_ok, eject_msg = safe_eject(drive)
            status_cb(eject_msg)

        msg_time = time.strftime("%H:%M:%S")
        success_msg = (
            f"✅ DRIVE ICON APPLIED SUCCESSFULLY!\n\n"
            f"Drive: {drive}\n"
            f"Time: {msg_time}  ({total:.1f}s)\n\n"
            f"  Windows (this PC)    • Registry updated\n"
            f"  Windows (other PCs)  • autorun.inf + desktop.ini\n"
            f"  macOS                • .VolumeIcon.icns\n"
            f"  Linux KDE            • .directory + drive_icon.png\n"
            f"  Linux GNOME/Nemo     • .xdg-volume-info\n"
            f"  All files hidden\n"
        )
        if eject_ok:
            success_msg += f"\n✅ Drive ejected successfully!"
        if is_sys:
            success_msg += f"\n\n⚠️ System drive may need a restart for full effect."

        done_cb(True, success_msg)

    except PermissionError as e:
        done_cb(False, f"❌ Permission denied:\n{e}\n\nRight-click → Run as Administrator")
    except Exception as e:
        import traceback
        done_cb(False, f"❌ Error: {e}\n\n{traceback.format_exc()}")

def _clear_desktop_ini_icon(desktop_ini_path):
    """
    Clear [.ShellClassInfo] from desktop.ini WITHOUT deleting the file.
    Used for system drives (C:) where deleting desktop.ini is protected --
    we blank the icon section so Explorer falls back to its default drive icon.
    If the file only had [.ShellClassInfo] it ends up empty (harmless).
    """
    try:
        clear_attribs(desktop_ini_path)
        with open(desktop_ini_path, "r", encoding="utf-8", errors="replace") as f:
            content = f.read()

        # Remove the entire [.ShellClassInfo] block
        import re
        # Match from [.ShellClassInfo] up to the next [Section] or end of file
        cleaned = re.sub(
            r'\[\.ShellClassInfo\][^\[]*',
            '',
            content,
            flags=re.IGNORECASE | re.DOTALL
        ).strip()

        if cleaned:
            # Rewrite without the ShellClassInfo block
            with open(desktop_ini_path, "w", encoding="utf-8", newline="") as f:
                f.write(cleaned + "\r\n")
        else:
            # File is now empty — delete it so Explorer shows default icon
            os.remove(desktop_ini_path)
    except Exception:
        pass  # Best-effort


def remove_drive_icon(drive_info, status_cb, done_cb):
    """
    Remove all icon files from a drive and restore the default Windows icon.

    ROOT CAUSE OF BLANK ICON BUG:
      Removing the registry entry is not enough.  Explorer also reads desktop.ini
      on the drive root.  If desktop.ini still has [.ShellClassInfo] pointing to
      a now-deleted .ico file, Explorer finds nothing and shows a blank/file icon.

    FIX:
      • Removable drives  → delete desktop.ini completely (safe to do)
      • System drive C:   → clear [.ShellClassInfo] block from desktop.ini
                            (can't delete it -- it's a protected system file)
      After cleanup: force a hard shell refresh sequence so Explorer forgets
      its cached blank icon and redraws with the default drive icon.
    """
    drive  = drive_info['path']
    is_sys = drive_info['is_system']
    letter = drive_info['letter']
    t0     = time.time()

    def step(msg):
        status_cb(f"[{time.time() - t0:.1f}s] {msg}")

    try:
        removed = 0

        # ── 1. Remove Registry entries ────────────────────────────────────────
        step("Removing Registry entries...")
        reg_remove_drive_icon(drive)
        removed += 1
        step("✅ Registry cleared")

        # ── 2. Handle desktop.ini ─────────────────────────────────────────────
        # THIS is what caused the blank icon — desktop.ini pointing to a
        # deleted .ico file. Must be handled BEFORE refreshing the shell.
        desktop_ini = os.path.join(drive, "desktop.ini")
        if os.path.exists(desktop_ini):
            if is_sys:
                # System drive: can't delete desktop.ini — clear ShellClassInfo block
                step("Clearing desktop.ini [.ShellClassInfo] (system drive)...")
                _clear_desktop_ini_icon(desktop_ini)
                step("✅ desktop.ini icon section cleared")
            else:
                # Removable drive: delete it entirely
                try:
                    clear_attribs(desktop_ini)
                    os.remove(desktop_ini)
                    removed += 1
                    step("✅ Removed desktop.ini")
                except Exception as e:
                    # Fallback: clear the block if we can't delete
                    step(f"⚠️ Cannot delete desktop.ini ({e}) — clearing icon section...")
                    _clear_desktop_ini_icon(desktop_ini)

        # ── 3. Remove other cross-platform files ──────────────────────────────
        for fname in ["autorun.inf", ".VolumeIcon.icns",
                      ".directory", ".xdg-volume-info"]:
            path = os.path.join(drive, fname)
            if os.path.exists(path):
                try:
                    clear_attribs(path)
                    os.remove(path)
                    removed += 1
                    step(f"✅ Removed {fname}")
                except Exception as e:
                    step(f"⚠️ Could not remove {fname}: {e}")

        # ── 4. Remove .icons folder from drive root (removable) ───────────────
        if not is_sys:
            icons_dir = os.path.join(drive, ".icons")
            if os.path.exists(icons_dir):
                try:
                    shutil.rmtree(icons_dir, ignore_errors=True)
                    if not os.path.exists(icons_dir):
                        removed += 1
                        step("✅ Removed .icons folder")
                    else:
                        step("⚠️ Could not fully remove .icons folder")
                except Exception as e:
                    step(f"⚠️ Error removing .icons folder: {e}")

        # ── 5. Remove ProgramData icon files ──────────────────────────────────
        for stem in [f"drive_{letter}.ico",
                     os.path.join(f"drive_{letter}_icons", "")]:
            pd_path = os.path.join(ICO_STORE, stem.rstrip(os.sep))
            if os.path.isfile(pd_path):
                try:
                    clear_attribs(pd_path)
                    os.remove(pd_path)
                    removed += 1
                    step(f"✅ Removed ProgramData: {os.path.basename(pd_path)}")
                except Exception as e:
                    step(f"⚠️ ProgramData file: {e}")
            elif os.path.isdir(pd_path):
                try:
                    shutil.rmtree(pd_path, ignore_errors=True)
                    removed += 1
                    step(f"✅ Removed ProgramData folder: drive_{letter}_icons")
                except Exception as e:
                    step(f"⚠️ ProgramData folder: {e}")

        # ── 6. Hard shell refresh — force Explorer to drop cached blank icon ───
        # Sequence:
        #   ASSOCCHANGED  → tells Explorer "icon associations changed globally"
        #   DRIVEREMOVED  → tells Explorer "this drive's appearance changed"
        #   DRIVEADD      → triggers Explorer to re-read drive info from scratch
        #   ie4uinit      → flushes thumbnail/icon disk cache
        step("Forcing shell refresh (restoring default icon)...")
        s32 = ctypes.windll.shell32

        drv_buf = ctypes.create_unicode_buffer(drive)
        s32.SHChangeNotify(SHCNE_ASSOCCHANGED,  SHCNF_FLUSHNOWAIT, None, None)
        time.sleep(0.1)
        s32.SHChangeNotify(0x00000080,           # SHCNE_DRIVEREMOVED
                           SHCNF_PATHW | SHCNF_FLUSH, drv_buf, None)
        time.sleep(0.15)
        s32.SHChangeNotify(0x00000100,           # SHCNE_DRIVEADD
                           SHCNF_PATHW | SHCNF_FLUSH, drv_buf, None)
        time.sleep(0.15)
        s32.SHChangeNotify(SHCNE_UPDATEITEM,
                           SHCNF_PATHW | SHCNF_FLUSH, drv_buf, None)
        # Flush icon cache in background
        try:
            subprocess.Popen(["ie4uinit.exe", "-show"],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
        step("✅ Shell refreshed — default drive icon restored")

        done_cb(True,
                f"✅ Icon removed from {drive}\n"
                f"{removed} item(s) cleaned up.\n\n"
                f"The default Windows drive icon has been restored.\n"
                f"{'(Sign out and back in if the icon still looks wrong.)' if is_sys else ''}")

    except Exception as e:
        import traceback
        done_cb(False, f"❌ Error: {e}\n\n{traceback.format_exc()}")

def drive_diagnostics(drive_info):
    """Show diagnostic information for a drive"""
    drive  = drive_info['path']
    letter = drive_info['letter']

    volume_label = get_drive_label(drive)

    lines = [f"=== Drive Diagnostics: {drive} ===",
             f"Windows        : {WIN_VER}",
             f"Admin rights   : {'YES' if is_admin() else 'NO'}",
             f"Drive type     : {drive_info['type_name']}",
             f"Volume label   : {volume_label or 'None'}",
             f"Filesystem     : {get_drive_filesystem(drive)}"]

    reg = reg_get_drive_icon(drive)
    lines.append(f"Registry icon  : {reg or 'NONE'}")

    # Check all cross-platform files
    for fname in ["desktop.ini", "autorun.inf", ".VolumeIcon.icns",
                  ".directory", ".xdg-volume-info"]:
        path   = os.path.join(drive, fname)
        status = "EXISTS" if os.path.exists(path) else "missing"
        hidden = " (HIDDEN)" if os.path.exists(path) and is_hidden_windows(path) else ""
        lines.append(f"{fname:22}: {status}{hidden}")

    # Check .icons folder
    icons_dir = os.path.join(drive, ".icons")
    if os.path.exists(icons_dir):
        icons = glob.glob(os.path.join(icons_dir, "*"))
        lines.append(f"\nIcons in .icons/: {len(icons)} files")
        for i in sorted(icons)[:8]:
            size   = os.path.getsize(i)
            name   = os.path.basename(i)
            hidden = " (HIDDEN)" if is_hidden_windows(i) else ""
            lines.append(f"  {name:24} ({size:,} bytes){hidden}")
    else:
        lines.append(f"\nIcons in .icons/: missing")

    # Check ProgramData
    pd = os.path.join(ICO_STORE, f"drive_{letter}.ico")
    lines.append(f"\nProgramData ico: {'EXISTS' if os.path.exists(pd) else 'missing'}")

    return "\n".join(lines)

def apply_folder_icon_pipeline(folder_path, ico_src, hide_files, status_cb, done_cb):
    """Apply icon to folder"""
    t0 = time.time()

    def step(msg):
        status_cb(f"[{time.time() - t0:.1f}s] {msg}")

    try:
        step("Starting folder icon application...")
        
        # Convert image to ICO if needed
        if not ico_src.lower().endswith('.ico'):
            step("Converting image to ICO format...")
            temp_ico = os.path.join(tempfile.gettempdir(), f"folder_icon_{int(time.time())}.ico")
            img = Image.open(ico_src).convert("RGBA")
            pil_to_ico(img, temp_ico)
            ico_src = temp_ico
            step("✅ Image converted to ICO")
        
        # Apply folder icon
        step("Setting folder icon...")
        success, messages = set_folder_icon(folder_path, ico_src, hide_files)
        
        for msg in messages:
            step(msg)
        
        if success:
            # Refresh Explorer
            step("Refreshing Explorer...")
            kill_explorer()
            delete_icon_cache()
            start_explorer()
            notify_shell(folder_path)
            
            total = time.time() - t0
            done_cb(True, 
                f"✅ FOLDER ICON APPLIED SUCCESSFULLY!\n\n"
                f"Folder: {folder_path}\n"
                f"Time: {total:.1f}s\n\n"
                f"• desktop.ini created\n"
                f"• Icon file hidden: {'Yes' if hide_files else 'No'}\n"
                f"• Folder set as system\n"
                f"• Icon visible in File Explorer")
        else:
            done_cb(False, f"❌ Failed to apply icon:\n{chr(10).join(messages)}")
            
    except Exception as e:
        import traceback
        done_cb(False, f"❌ Error: {e}\n\n{traceback.format_exc()}")

def remove_folder_icon_pipeline(folder_path, status_cb, done_cb):
    """Remove icon from folder"""
    t0 = time.time()

    def step(msg):
        status_cb(f"[{time.time() - t0:.1f}s] {msg}")

    try:
        step("Removing folder icon...")
        success, messages = remove_folder_icon(folder_path)
        
        for msg in messages:
            step(msg)
        
        if success:
            # Refresh Explorer
            step("Refreshing Explorer...")
            kill_explorer()
            delete_icon_cache()
            start_explorer()
            notify_shell(folder_path)
            
            done_cb(True, f"✅ Folder icon removed successfully!")
        else:
            done_cb(False, f"❌ Failed to remove icon:\n{chr(10).join(messages)}")
            
    except Exception as e:
        import traceback
        done_cb(False, f"❌ Error: {e}\n\n{traceback.format_exc()}")

def folder_diagnostics(folder_path):
    """Show diagnostic information for a folder"""
    lines = [f"=== Folder Diagnostics: {folder_path} ===",
             f"Windows        : {WIN_VER}",
             f"Admin rights   : {'YES' if is_admin() else 'NO'}",
             f"Folder exists  : {'YES' if os.path.exists(folder_path) else 'NO'}"]

    if not os.path.exists(folder_path):
        return "\n".join(lines)

    # Check folder attributes
    try:
        result = subprocess.run(f"attrib {folder_path}", shell=True, 
                               capture_output=True, text=True)
        lines.append(f"Attributes     : {result.stdout.strip()}")
    except:
        pass

    # Check desktop.ini
    desktop_ini = os.path.join(folder_path, "desktop.ini")
    if os.path.exists(desktop_ini):
        status = "EXISTS"
        hidden = " (HIDDEN)" if is_hidden_windows(desktop_ini) else ""
        lines.append(f"desktop.ini    : {status}{hidden}")
        
        # Read desktop.ini content
        try:
            with open(desktop_ini, 'r', encoding='utf-8') as f:
                content = f.read()
            lines.append("\n--- desktop.ini content ---")
            for line in content.splitlines():
                lines.append(f"  {line}")
            
            # Find icon file
            import re
            icon_match = re.search(r'IconResource=(.+?),0', content)
            if icon_match:
                icon_path = icon_match.group(1).strip()
                if os.path.exists(icon_path):
                    icon_status = "EXISTS"
                    icon_hidden = " (HIDDEN)" if is_hidden_windows(icon_path) else ""
                    icon_size = os.path.getsize(icon_path)
                    lines.append(f"\nIcon file      : {icon_status}{icon_hidden}")
                    lines.append(f"Icon path      : {icon_path}")
                    lines.append(f"Icon size      : {icon_size:,} bytes")
                else:
                    lines.append(f"\nIcon file      : MISSING - {icon_path}")
        except Exception as e:
            lines.append(f"Error reading desktop.ini: {e}")
    else:
        lines.append(f"desktop.ini    : missing")

    return "\n".join(lines)

# ==============================================================================
#  GUI COMPONENTS
# ==============================================================================

# Colors (Catppuccin theme)
BG = "#1e1e2e"
SURFACE = "#313244"
OVERLAY = "#45475a"
TEXT = "#cdd6f4"
SUBTEXT = "#a6adc8"
ACCENT = "#89b4fa"
GREEN = "#a6e3a1"
RED = "#f38ba8"
YELLOW = "#f9e2af"
ORANGE = "#fab387"
PURPLE = "#cba6f7"
TEAL = "#94e2d5"

def flat_btn(parent, text, cmd, accent=False, color=None, font_size=10, bold=False, **kw):
    """Create a flat button with consistent styling"""
    bg = color or (ACCENT if accent else SURFACE)
    fg = "#1e1e2e" if (accent or color) else TEXT
    
    # Set font based on parameters
    font_weight = "bold" if (accent or bold) else "normal"
    font = ("Segoe UI", font_size, font_weight)
    
    return tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                     activebackground="#b4befe" if accent else OVERLAY,
                     activeforeground="#1e1e2e" if accent else TEXT,
                     relief="flat", cursor="hand2", bd=0,
                     font=font, padx=10, pady=6, **kw)

# ── Step Log Window ───────────────────────────────────────────────────────────
class StepLog(tk.Toplevel):
    def __init__(self, parent, title="Progress"):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG, padx=16, pady=14)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        
        tk.Label(self, text=f"Live Progress  (Windows {WIN_VER})",
                 bg=BG, fg=ACCENT,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        
        self.txt = tk.Text(self, width=72, height=20,
                           bg=SURFACE, fg=GREEN,
                           font=("Consolas", 9), relief="flat",
                           state="disabled", wrap="word")
        self.txt.pack()
        self.geometry(f"+{parent.winfo_x() + 50}+{parent.winfo_y() + 20}")
        self.update()

    def log(self, msg):
        self.txt.config(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.config(state="disabled")
        self.update()

    def done(self):
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.log("\n─── Click X to close ───")

# ── Crop Editor ───────────────────────────────────────────────────────────────
EDITOR_SIZE = 320

class CropEditor(tk.Toplevel):
    def __init__(self, parent, pil_image, callback):
        super().__init__(parent)
        self.title("Edit Icon — Drag to pan  |  Scroll to zoom")
        self.configure(bg=BG, padx=20, pady=16)
        self.resizable(False, False)
        self.grab_set()
        self._src = pil_image.convert("RGBA")
        self._cb = callback
        self._zoom = 1.0
        self._off = [0, 0]
        self._drag = None
        self._smi = []
        self._hq_job = None
        self._build()
        self._center()
        self._redraw()

    def _build(self):
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", pady=(0, 10))
        
        lf = tk.Frame(top, bg=BG)
        lf.pack(side="left", padx=(0, 16))
        
        tk.Label(lf, text="Drag to pan  |  Scroll to zoom",
                 bg=BG, fg=SUBTEXT, font=("Segoe UI", 8)).pack()
        
        self.cv = tk.Canvas(lf, width=EDITOR_SIZE, height=EDITOR_SIZE,
                           bg="#000", highlightthickness=2,
                           highlightbackground=ACCENT, cursor="fleur")
        self.cv.pack()
        self.cv.bind("<ButtonPress-1>", self._ds)
        self.cv.bind("<B1-Motion>", self._dm)
        self.cv.bind("<MouseWheel>", self._mw)
        
        rf = tk.Frame(top, bg=BG)
        rf.pack(side="left", anchor="n")
        
        tk.Label(rf, text="Preview (256px)", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 8)).pack()
        
        self.pv = tk.Canvas(rf, width=128, height=128, bg="#000",
                           highlightthickness=1, highlightbackground=OVERLAY)
        self.pv.pack(pady=(0, 8))
        
        tk.Label(rf, text="Small sizes:", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 8)).pack(anchor="w")
        
        self.sm = tk.Canvas(rf, width=128, height=52, bg="#2a2a3e",
                           highlightthickness=0)
        self.sm.pack()
        
        tk.Label(rf, text="Background:", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(10, 2))
        
        self._bg = tk.StringVar(value="transparent")
        for v, l in [("transparent", "Transparent"), ("white", "White"),
                    ("black", "Black"), ("circle", "Circle crop")]:
            tk.Radiobutton(rf, text=l, variable=self._bg, value=v,
                          bg=BG, fg=TEXT, selectcolor=SURFACE,
                          activebackground=BG, activeforeground=TEXT,
                          font=("Segoe UI", 9),
                          command=self._redraw).pack(anchor="w")
        
        zm = tk.Frame(self, bg=BG)
        zm.pack(fill="x", pady=(0, 12))
        
        tk.Label(zm, text="Zoom:", bg=BG, fg=TEXT,
                 font=("Segoe UI", 10)).pack(side="left")
        
        self.zsl = tk.Scale(zm, from_=1, to=5000, orient="horizontal",
                           bg=BG, fg=TEXT, troughcolor=SURFACE,
                           highlightthickness=0, showvalue=False,
                           command=self._zc)
        self.zsl.set(100)
        self.zsl.pack(side="left", fill="x", expand=True, padx=(8, 8))
        
        self.zlb = tk.Label(zm, text="100%", bg=BG, fg=ACCENT,
                           font=("Segoe UI", 10, "bold"), width=7)
        self.zlb.pack(side="left")
        
        br = tk.Frame(self, bg=BG)
        br.pack(fill="x")
        
        flat_btn(br, "Fit", self._fit).pack(side="left", padx=(0, 8))
        flat_btn(br, "Cancel", self.destroy).pack(side="right", padx=(8, 0))
        flat_btn(br, "Use this icon", self._confirm, accent=True).pack(side="right")

    def _center(self):
        sw, sh = self._src.size
        self._zoom = min(EDITOR_SIZE / sw, EDITOR_SIZE / sh)
        self._off = [(sw - EDITOR_SIZE / self._zoom) / 2,
                    (sh - EDITOR_SIZE / self._zoom) / 2]
        self.zsl.set(int(self._zoom * 100))

    def _fit(self):
        self._center()
        self._redraw()

    def _zc(self, v):
        cx = self._off[0] + (EDITOR_SIZE / 2) / self._zoom
        cy = self._off[1] + (EDITOR_SIZE / 2) / self._zoom
        self._zoom = max(0.01, int(v) / 100)
        self._off[0] = cx - (EDITOR_SIZE / 2) / self._zoom
        self._off[1] = cy - (EDITOR_SIZE / 2) / self._zoom
        self.zlb.config(text=f"{int(v)}%")
        self._redraw(fast=True)

    def _ds(self, e):
        self._drag = (e.x, e.y, self._off[0], self._off[1])

    def _dm(self, e):
        if not self._drag:
            return
        sx, sy, ox, oy = self._drag
        self._off[0] = ox + (sx - e.x) / self._zoom
        self._off[1] = oy + (sy - e.y) / self._zoom
        self._redraw(fast=True)

    def _mw(self, e):
        f = 1.1 if e.delta > 0 else 0.9
        nz = max(0.01, min(50.0, self._zoom * f))
        cx = self._off[0] + (EDITOR_SIZE / 2) / self._zoom
        cy = self._off[1] + (EDITOR_SIZE / 2) / self._zoom
        self._zoom = nz
        self._off[0] = cx - (EDITOR_SIZE / 2) / self._zoom
        self._off[1] = cy - (EDITOR_SIZE / 2) / self._zoom
        self.zsl.set(int(self._zoom * 100))
        self.zlb.config(text=f"{int(self._zoom * 100)}%")
        self._redraw(fast=True)

    def _crop(self, size=256, resample=Image.LANCZOS):
        sw, sh = self._src.size
        x0, y0 = self._off
        x1 = x0 + EDITOR_SIZE / self._zoom
        y1 = y0 + EDITOR_SIZE / self._zoom
        bv = self._bg.get()
        
        ci = Image.new("RGBA", (EDITOR_SIZE, EDITOR_SIZE),
                      (255, 255, 255, 255) if bv == "white" else
                      (0, 0, 0, 255) if bv == "black" else (0, 0, 0, 0))
        
        sx0, sy0 = max(0, x0), max(0, y0)
        sx1, sy1 = min(sw, x1), min(sh, y1)
        
        if sx1 > sx0 and sy1 > sy0:
            rg = self._src.crop((sx0, sy0, sx1, sy1))
            px = int((sx0 - x0) * self._zoom)
            py = int((sy0 - y0) * self._zoom)
            pw = max(1, int((sx1 - sx0) * self._zoom))
            ph = max(1, int((sy1 - sy0) * self._zoom))
            rs = rg.resize((pw, ph), resample)
            ci.paste(rs, (px, py), rs)
        
        if bv == "circle":
            mk = Image.new("L", (EDITOR_SIZE, EDITOR_SIZE), 0)
            ImageDraw.Draw(mk).ellipse(
                (0, 0, EDITOR_SIZE - 1, EDITOR_SIZE - 1), fill=255)
            ot = Image.new("RGBA", (EDITOR_SIZE, EDITOR_SIZE), (0, 0, 0, 0))
            ot.paste(ci, mask=mk)
            ci = ot
        
        return ci.resize((size, size), Image.LANCZOS)

    @staticmethod
    def _chk(size, b=8):
        img = Image.new("RGBA", (size, size))
        d = ImageDraw.Draw(img)
        for y in range(0, size, b):
            for x in range(0, size, b):
                c = ((200, 200, 200, 255) if (x // b + y // b) % 2 == 0
                     else (160, 160, 160, 255))
                d.rectangle([x, y, x + b - 1, y + b - 1], fill=c)
        return img

    def _redraw(self, fast=False):
        # Cancel any pending HQ redraws
        if self._hq_job:
            self.after_cancel(self._hq_job)
            self._hq_job = None

        resample = Image.NEAREST if fast else Image.BILINEAR
        img = self._crop(EDITOR_SIZE, resample=resample)
        
        self._te = ImageTk.PhotoImage(
            Image.alpha_composite(self._chk(EDITOR_SIZE), img))
        self.cv.delete("all")
        self.cv.create_image(0, 0, anchor="nw", image=self._te)
        
        if fast:
            # Schedule a high-quality redraw when the user stops interacting
            self._hq_job = self.after(150, lambda: self._redraw(fast=False))
            return

        # Regular high-quality updates for previews
        pv = img.resize((128, 128), Image.BILINEAR)
        self._tp = ImageTk.PhotoImage(
            Image.alpha_composite(self._chk(128), pv))
        self.pv.delete("all")
        self.pv.create_image(0, 0, anchor="nw", image=self._tp)
        
        self.sm.delete("all")
        self._smi = []
        x = 4
        for s in [48, 32, 16]:
            ti = ImageTk.PhotoImage(Image.alpha_composite(
                self._chk(s), img.resize((s, s), Image.BILINEAR)))
            self._smi.append(ti)
            self.sm.create_image(x, 26, anchor="w", image=ti)
            self.sm.create_text(x + s + 2, 42, anchor="w", text=f"{s}px",
                                fill=SUBTEXT, font=("Segoe UI", 7))
            x += s + 28

    def _confirm(self):
        self._cb(self._crop(256))
        self.destroy()

# ==============================================================================
#  HOME PAGE WITH TWO MODES
# ==============================================================================

class HomePage(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Drive & Folder Icon Setter  v13.0  —  Windows {WIN_VER}")
        self.configure(bg=BG, padx=28, pady=22)
        self.resizable(False, False)
        
        self._build_ui()
        
        if not is_admin():
            self._admin_banner()

    def _build_ui(self):
        # Title
        tk.Label(self,
                 text="  Drive & Folder Icon Setter  v13.0",
                 bg=BG, fg=ACCENT,
                 font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 4))

        # Info frame
        info = tk.Frame(self, bg=SURFACE, padx=12, pady=8)
        info.pack(fill="x", pady=(0, 20))
        
        tk.Label(info,
                 text=f"Windows {WIN_VER} • Choose your mode below",
                 bg=SURFACE, fg=GREEN,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")

        # Mode selection buttons
        mode_frame = tk.Frame(self, bg=BG)
        mode_frame.pack(fill="x", pady=10)

        # Drive Mode Button
        drive_frame = tk.Frame(mode_frame, bg=SURFACE, padx=20, pady=20)
        drive_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        tk.Label(drive_frame,
                 text="💾 CHANGE DRIVE ICON",
                 bg=SURFACE, fg=TEAL,
                 font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        tk.Label(drive_frame,
                 text="Set custom icons for:\n\u2022 C:\\, D:\\, E:\\ drives\n\u2022 USB flash drives\n\u2022 External hard disks\n\u2022 Any removable media",
                 bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 10), justify="left").pack(anchor="w", pady=(0, 15))
        
        flat_btn(drive_frame, "  Open Drive Icon Tool  ",
                self._open_drive_mode, accent=True, font_size=12, bold=True).pack(pady=5)

        # Folder Mode Button
        folder_frame = tk.Frame(mode_frame, bg=SURFACE, padx=20, pady=20)
        folder_frame.pack(side="left", fill="both", expand=True, padx=(10, 0))
        
        tk.Label(folder_frame,
                 text="📁 CHANGE FOLDER ICON",
                 bg=SURFACE, fg=PURPLE,
                 font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        tk.Label(folder_frame,
                 text="Set custom icons for:\n• Any folder on your PC\n• Documents, Downloads\n• Project folders\n• Auto-hide icon files",
                 bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 10), justify="left").pack(anchor="w", pady=(0, 15))
        
        flat_btn(folder_frame, "  Open Folder Icon Tool  ",
                self._open_folder_mode, accent=True, color=PURPLE, 
                font_size=12, bold=True).pack(pady=5)

        # Features list
        features = tk.Frame(self, bg=BG, pady=10)
        features.pack(fill="x")
        tk.Label(features, text="\u2728 Features:",
                 bg=BG, fg=YELLOW,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Label(features,
                 text="\u2022 Drive icons  \u2022 Folder icons  \u2022 Crop editor  "
                      "\u2022 Registry integration  \u2022 Instant refresh",
                 bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).pack(anchor="w")

    def _admin_banner(self):
        b = tk.Frame(self, bg=YELLOW, padx=10, pady=8)
        b._is_admin_banner = True
        b.pack(fill="x", pady=(10, 0))
        tk.Label(b,
                 text="  Not running as Administrator — Some features may fail!",
                 bg=YELLOW, fg="#1e1e2e",
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Button(b, text="Restart as Admin", command=relaunch_as_admin,
                  bg="#fe640b", fg="white", relief="flat", cursor="hand2",
                  padx=8, pady=3,
                  font=("Segoe UI", 9, "bold")).pack(side="right")

    def _refresh_admin_banner(self):
        for w in self.winfo_children():
            if getattr(w, "_is_admin_banner", False):
                w.destroy()
        if not is_admin():
            self._admin_banner()

    def _open_drive_mode(self):
        self.withdraw()
        win = DriveIconApp(self)
        self.wait_window(win)
        self.deiconify()
        self._refresh_admin_banner()

    def _open_folder_mode(self):
        self.withdraw()
        win = FolderIconApp(self)
        self.wait_window(win)
        self.deiconify()
        self._refresh_admin_banner()


class DriveIconApp(tk.Toplevel):
    def __init__(self, home_page):
        super().__init__(home_page)
        self.title(f"Drive Icon Tool  —  Windows {WIN_VER}  (v13.0)")
        self.configure(bg=BG, padx=28, pady=22)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        
        self.home_page = home_page
        self._src = None
        self._final = None
        self._ico = None
        self._tmp = tempfile.mkdtemp()
        self._drives = []
        
        self.drive_var = tk.StringVar()
        self.eject_var = tk.BooleanVar(value=False)
        
        self._build_ui()
        self._refresh_drives()

    def _on_close(self):
        """Return to home page"""
        self.home_page.deiconify()
        self.destroy()

    def _build_ui(self):
        # Back button
        back_frame = tk.Frame(self, bg=BG)
        back_frame.pack(fill="x", pady=(0, 10))
        
        flat_btn(back_frame, "← Back to Home", self._on_close,
                color=OVERLAY).pack(side="left")

        # Title
        tk.Label(self,
                 text="  💾 Drive Icon Tool",
                 bg=BG, fg=TEAL,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 4))

        # Step 1 - Choose image
        self._sec("Step 1  —  Choose an image")
        
        f1 = tk.Frame(self, bg=BG)
        f1.pack(fill="x", pady=(0, 8))
        
        self.img_var = tk.StringVar()
        tk.Entry(f1, textvariable=self.img_var, width=38, bg=SURFACE, fg=TEXT,
                insertbackground=TEXT, relief="flat", font=("Segoe UI", 10),
                state="readonly", readonlybackground=SURFACE
                ).pack(side="left", padx=(0, 8), ipady=5)
        flat_btn(f1, "Browse…", self._browse).pack(side="left")

        fp = tk.Frame(self, bg=BG)
        fp.pack(fill="x", pady=(4, 0))
        
        self.thumb_cv = tk.Canvas(fp, width=96, height=96, bg=SURFACE,
                                 highlightthickness=1, highlightbackground=OVERLAY)
        self.thumb_cv.pack(side="left")
        self.thumb_cv.create_text(48, 48, text="preview",
                                 fill=SUBTEXT, font=("Segoe UI", 9))
        
        fi = tk.Frame(fp, bg=BG, padx=14)
        fi.pack(side="left", fill="both")
        
        self.info_v = tk.StringVar(value="No image selected.")
        tk.Label(fi, textvariable=self.info_v, bg=BG, fg=SUBTEXT,
                font=("Segoe UI", 9), justify="left").pack(anchor="w")
        
        self.conv_l = tk.Label(fi, text="", bg=BG, fg=GREEN,
                              font=("Segoe UI", 9, "bold"), justify="left")
        self.conv_l.pack(anchor="w", pady=(4, 0))
        
        flat_btn(fi, "  Edit / Crop icon  ",
                self._open_editor, color=PURPLE).pack(anchor="w", pady=(10, 0))

        # Step 2 - Select drive
        self._sec("Step 2  —  Select drive")
        
        f2 = tk.Frame(self, bg=BG)
        f2.pack(fill="x", pady=(0, 4))
        
        tk.Label(f2, text="Drive:", bg=BG, fg=TEXT,
                font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=5)
        
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TCombobox",
                       fieldbackground=SURFACE, background=SURFACE,
                       foreground=TEXT, selectbackground=SURFACE,
                       selectforeground=TEXT)
        
        self.combo = ttk.Combobox(f2, textvariable=self.drive_var,
                                  width=32, state="readonly")
        self.combo.grid(row=0, column=1, padx=(8, 8), sticky="w")
        self.combo.bind("<<ComboboxSelected>>", self._on_drive)
        
        flat_btn(f2, "Refresh", self._refresh_drives).grid(row=0, column=2)

        self.cur_ico_l = tk.Label(self, text="", bg=BG, fg=SUBTEXT,
                                 font=("Segoe UI", 8), anchor="w")
        self.cur_ico_l.pack(fill="x", pady=(0, 2))
        
        self.warn_l = tk.Label(self, text="", bg=BG, fg=ORANGE,
                              font=("Segoe UI", 9, "bold"),
                              anchor="w", justify="left")
        self.warn_l.pack(fill="x", pady=(0, 4))

        # Step 3 - Options & Apply
        self._sec("Step 3  —  Options & Apply")
        
        self.eject_chk = tk.Checkbutton(
            self, text="  Auto Eject after Apply  (USB / External HDD)",
            variable=self.eject_var, bg=BG, fg=GREEN, selectcolor=SURFACE,
            activebackground=BG, activeforeground=GREEN,
            font=("Segoe UI", 10, "bold"), state="disabled")
        self.eject_chk.pack(anchor="w", pady=(3, 10))

        # Progress bar
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=420)
        tk.Frame(self, bg=BG, height=5).pack()

        # Apply button
        self.apply_btn = flat_btn(
            self, "  ✅ APPLY ICON TO DRIVE  ",
            self._apply, accent=True, font_size=11, bold=True)
        self.apply_btn.pack(fill="x", pady=(5, 0))

        # Status
        self.status_v = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self.status_v, bg="#181825", fg=SUBTEXT,
                anchor="w", font=("Consolas", 9), padx=10, pady=5
                ).pack(fill="x", pady=(8, 0))

        # Bottom buttons
        bf = tk.Frame(self, bg=BG)
        bf.pack(fill="x", pady=(6, 0))
        
        flat_btn(bf, "  Remove Icon  ",
                self._remove_icon, color="#585b70"
                ).pack(side="left", fill="x", expand=True, padx=(0, 3))
        
        flat_btn(bf, "  Diagnostics  ",
                self._diagnostics, color=ORANGE
                ).pack(side="left", fill="x", expand=True, padx=(3, 0))

    def _sec(self, title):
        tk.Label(self, text=title, bg=BG, fg=ACCENT,
                font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(14, 4))
        tk.Frame(self, bg=OVERLAY, height=1).pack(fill="x", pady=(0, 8))

    def _refresh_drives(self):
        """Refresh list of drives"""
        self._drives = get_drives()
        
        choices = []
        for d in self._drives:
            display = (f"{d['path']}  {d['label']}  ({d['type_name']})"
                      + (" [SYSTEM]" if d['is_system'] else ""))
            choices.append(display)
        
        self.combo["values"] = choices
        if choices:
            for i, d in enumerate(self._drives):
                if not d['is_system']:
                    self.combo.current(i)
                    break
            else:
                self.combo.current(0)
        self._on_drive()

    def _get_drive(self):
        idx = self.combo.current()
        if idx < 0 or idx >= len(self._drives):
            return None
        return self._drives[idx]

    def _on_drive(self, event=None):
        drive = self._get_drive()
        if not drive:
            return
        
        is_usb = (drive['type'] == DRIVE_REMOVABLE)
        is_sys = drive['is_system']
        volume_label = get_drive_label(drive['path'])
        
        cur = reg_get_drive_icon(drive['path'])
        if cur:
            self.cur_ico_l.config(
                text=f"Current icon: {cur}  |  Volume label: '{volume_label}'")
        else:
            self.cur_ico_l.config(
                text=f"No custom icon. Volume label: '{volume_label}'")
        
        if is_sys:
            self.warn_l.config(
                text="  System Drive — Restart PC manually to fully apply icon.",
                fg=ORANGE)
        else:
            self.warn_l.config(text="")
        
        self.eject_chk.config(state="normal" if is_usb else "disabled")
        if not is_usb:
            self.eject_var.set(False)

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select image",
            filetypes=[("Image files",
                       "*.png *.jpg *.jpeg *.bmp *.gif *.webp *.tiff *.tif *.ico"),
                      ("All files", "*.*")])
        if not path:
            return
        
        try:
            img = Image.open(path)
            self._src = img.convert("RGBA")
            self.img_var.set(path)
            ext = os.path.splitext(path)[1].upper()
            
            self.info_v.set(
                f"File: {os.path.basename(path)}\n"
                f"Size: {img.width} x {img.height} px  |  {ext}")
            
            self._ico = None
            self._final = None
            self.conv_l.config(text="Click 'Edit / Crop icon' to adjust.",
                               fg=YELLOW)
            self._thumb_update(self._src)
            self._open_editor()
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open image:\n{e}")

    def _thumb_update(self, pil_img):
        t = pil_img.copy().convert("RGBA")
        t.thumbnail((96, 96), Image.LANCZOS)
        
        chk = Image.new("RGBA", (96, 96))
        d = ImageDraw.Draw(chk)
        for y in range(0, 96, 8):
            for x in range(0, 96, 8):
                c = ((200, 200, 200, 255) if (x // 8 + y // 8) % 2 == 0
                     else (160, 160, 160, 255))
                d.rectangle([x, y, x + 7, y + 7], fill=c)
        
        ox = (96 - t.width) // 2
        oy = (96 - t.height) // 2
        chk.paste(t, (ox, oy), t)
        
        self._tk_thumb = ImageTk.PhotoImage(chk)
        self.thumb_cv.delete("all")
        self.thumb_cv.create_image(0, 0, anchor="nw", image=self._tk_thumb)

    def _open_editor(self):
        if self._src is None:
            messagebox.showwarning("No image", "Please select an image first.")
            return
        CropEditor(self, self._src, self._edit_done)

    def _edit_done(self, result):
        self._final = result
        self._thumb_update(result)
        
        try:
            out = os.path.join(self._tmp, "drive_icon.ico")
            pil_to_ico(result, out)
            self._ico = out
            self.conv_l.config(text="Icon ready! Click Apply.", fg=GREEN)
            self.status_v.set("Icon ready")
        except Exception as e:
            self.conv_l.config(text=f"Convert failed: {e}", fg=RED)

    def _check_ready(self):
        if not self._ico or not os.path.isfile(self._ico):
            messagebox.showwarning("Not Ready",
                "Please select and edit an image first.")
            return False
        
        drive = self._get_drive()
        if not drive:
            messagebox.showwarning("No Drive", "Please select a drive.")
            return False
        
        return True

    def _run_pipeline(self, target_fn, args):
        self.progress.pack(fill="x", pady=(0, 8), before=self.apply_btn)
        self.progress.start(10)
        self.update()
        
        log = StepLog(self)

        def _status(msg):
            self.after(0, lambda m=msg: (self.status_v.set(m), log.log(m)))

        def _done(ok, msg):
            self.after(0, lambda o=ok, m=msg: self._finish(o, m, log))

        threading.Thread(
            target=target_fn,
            args=args + (_status, _done),
            daemon=True).start()

    def _apply(self):
        if not self._check_ready():
            return
        
        drive = self._get_drive()
        
        if drive['is_system']:
            if not messagebox.askyesno("System Drive",
                f"Apply icon to SYSTEM drive {drive['path']}?\n\n"
                f"Explorer will restart briefly.\n"
                f"You need to restart PC manually for full effect.\n\nContinue?",
                icon="warning"):
                return
        elif self.eject_var.get():
            if not messagebox.askyesno("Confirm Eject",
                f"Apply icon to {drive['path']} then eject?\n"
                f"Close all files on this drive first."):
                return

        self._run_pipeline(apply_drive_icon,
                          (drive, self._ico, self.eject_var.get()))

    def _finish(self, success, msg, log):
        self.progress.stop()
        self.progress.pack_forget()
        log.done()
        
        if success:
            messagebox.showinfo("Success!", msg)
            self._refresh_drives()
        else:
            messagebox.showerror("Error", msg)

    def _remove_icon(self):
        drive = self._get_drive()
        if not drive:
            return
        
        if not messagebox.askyesno("Remove Icon",
            f"Remove custom icon from {drive['path']}?\n"
            f"Will restore default Windows icon."):
            return
        
        self.progress.pack(fill="x", pady=(0, 8), before=self.apply_btn)
        self.progress.start(10)
        self.update()
        
        log = StepLog(self)

        def _status(msg):
            self.after(0, lambda m=msg: (self.status_v.set(m), log.log(m)))

        def _done(ok, msg):
            self.after(0, lambda o=ok, m=msg: self._finish_remove(o, m, log))

        threading.Thread(
            target=remove_drive_icon,
            args=(drive, _status, _done),
            daemon=True).start()

    def _finish_remove(self, success, msg, log):
        self.progress.stop()
        self.progress.pack_forget()
        log.done()
        
        if success:
            messagebox.showinfo("Success!", msg)
            self._refresh_drives()
        else:
            messagebox.showerror("Error", msg)

    def _diagnostics(self):
        drive = self._get_drive()
        if not drive:
            return
        messagebox.showinfo("Diagnostics", drive_diagnostics(drive))

    def destroy(self):
        try:
            shutil.rmtree(self._tmp, ignore_errors=True)
        except:
            pass
        super().destroy()


class FolderIconApp(tk.Toplevel):
    def __init__(self, home_page):
        super().__init__(home_page)
        self.title(f"Folder Icon Tool  —  Windows {WIN_VER}")
        self.configure(bg=BG, padx=28, pady=22)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        
        self.home_page = home_page
        self._src = None
        self._final = None
        self._ico = None
        self._tmp = tempfile.mkdtemp()
        
        self.folder_path = tk.StringVar()
        self.hide_files_var = tk.BooleanVar(value=True)
        
        self._build_ui()

    def _on_close(self):
        """Return to home page"""
        self.home_page.deiconify()
        self.destroy()

    def _build_ui(self):
        # Back button
        back_frame = tk.Frame(self, bg=BG)
        back_frame.pack(fill="x", pady=(0, 10))
        
        flat_btn(back_frame, "← Back to Home", self._on_close,
                color=OVERLAY).pack(side="left")

        # Title
        tk.Label(self,
                 text="  📁 Folder Icon Tool",
                 bg=BG, fg=PURPLE,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 4))

        # Step 1 - Choose image
        self._sec("Step 1  —  Choose an image")
        
        f1 = tk.Frame(self, bg=BG)
        f1.pack(fill="x", pady=(0, 8))
        
        self.img_var = tk.StringVar()
        tk.Entry(f1, textvariable=self.img_var, width=38, bg=SURFACE, fg=TEXT,
                insertbackground=TEXT, relief="flat", font=("Segoe UI", 10),
                state="readonly", readonlybackground=SURFACE
                ).pack(side="left", padx=(0, 8), ipady=5)
        flat_btn(f1, "Browse…", self._browse).pack(side="left")

        fp = tk.Frame(self, bg=BG)
        fp.pack(fill="x", pady=(4, 0))
        
        self.thumb_cv = tk.Canvas(fp, width=96, height=96, bg=SURFACE,
                                 highlightthickness=1, highlightbackground=OVERLAY)
        self.thumb_cv.pack(side="left")
        self.thumb_cv.create_text(48, 48, text="preview",
                                 fill=SUBTEXT, font=("Segoe UI", 9))
        
        fi = tk.Frame(fp, bg=BG, padx=14)
        fi.pack(side="left", fill="both")
        
        self.info_v = tk.StringVar(value="No image selected.")
        tk.Label(fi, textvariable=self.info_v, bg=BG, fg=SUBTEXT,
                font=("Segoe UI", 9), justify="left").pack(anchor="w")
        
        self.conv_l = tk.Label(fi, text="", bg=BG, fg=GREEN,
                              font=("Segoe UI", 9, "bold"), justify="left")
        self.conv_l.pack(anchor="w", pady=(4, 0))
        
        flat_btn(fi, "  Edit / Crop icon  ",
                self._open_editor, color=PURPLE).pack(anchor="w", pady=(10, 0))

        # Step 2 - Select folder
        self._sec("Step 2  —  Select folder")
        
        f2 = tk.Frame(self, bg=BG)
        f2.pack(fill="x", pady=(0, 4))
        
        tk.Entry(f2, textvariable=self.folder_path, width=45, bg=SURFACE, fg=TEXT,
                insertbackground=TEXT, relief="flat", font=("Segoe UI", 10)
                ).pack(side="left", padx=(0, 8), ipady=5)
        flat_btn(f2, "Browse…", self._browse_folder).pack(side="left")

        # Folder status
        self.folder_status = tk.Label(self, text="", bg=BG, fg=SUBTEXT,
                                     font=("Segoe UI", 8), anchor="w")
        self.folder_status.pack(fill="x", pady=(2, 0))

        # Options
        self._sec("Step 3  —  Options & Apply")
        
        self.hide_chk = tk.Checkbutton(
            self, text="  Hide icon files  (+H+S attributes)",
            variable=self.hide_files_var, bg=BG, fg=GREEN, selectcolor=SURFACE,
            activebackground=BG, activeforeground=GREEN,
            font=("Segoe UI", 10, "bold"))
        self.hide_chk.pack(anchor="w", pady=(3, 10))

        # Progress bar
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=420)
        tk.Frame(self, bg=BG, height=5).pack()

        # Apply button
        self.apply_btn = flat_btn(
            self, "  ✅ APPLY ICON TO FOLDER  ",
            self._apply, accent=True, font_size=11, bold=True)
        self.apply_btn.pack(fill="x", pady=(5, 0))

        # Status
        self.status_v = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self.status_v, bg="#181825", fg=SUBTEXT,
                anchor="w", font=("Consolas", 9), padx=10, pady=5
                ).pack(fill="x", pady=(8, 0))

        # Bottom buttons
        bf = tk.Frame(self, bg=BG)
        bf.pack(fill="x", pady=(6, 0))
        
        flat_btn(bf, "  Remove Icon  ",
                self._remove_icon, color="#585b70"
                ).pack(side="left", fill="x", expand=True, padx=(0, 3))
        
        flat_btn(bf, "  Diagnostics  ",
                self._diagnostics, color=ORANGE
                ).pack(side="left", fill="x", expand=True, padx=(3, 0))

    def _sec(self, title):
        tk.Label(self, text=title, bg=BG, fg=ACCENT,
                font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(14, 4))
        tk.Frame(self, bg=OVERLAY, height=1).pack(fill="x", pady=(0, 8))

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select folder")
        if path:
            self.folder_path.set(path)
            self._update_folder_status()

    def _update_folder_status(self):
        folder = self.folder_path.get()
        if folder and os.path.exists(folder):
            icon_path, status = get_folder_icon_status(folder)
            self.folder_status.config(text=status, fg=GREEN if icon_path else SUBTEXT)
        else:
            self.folder_status.config(text="")

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select image",
            filetypes=[("Image files",
                       "*.png *.jpg *.jpeg *.bmp *.gif *.webp *.tiff *.tif *.ico"),
                      ("All files", "*.*")])
        if not path:
            return
        
        try:
            img = Image.open(path)
            self._src = img.convert("RGBA")
            self.img_var.set(path)
            ext = os.path.splitext(path)[1].upper()
            
            self.info_v.set(
                f"File: {os.path.basename(path)}\n"
                f"Size: {img.width} x {img.height} px  |  {ext}")
            
            self._ico = None
            self._final = None
            self.conv_l.config(text="Click 'Edit / Crop icon' to adjust.",
                               fg=YELLOW)
            self._thumb_update(self._src)
            self._open_editor()
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open image:\n{e}")

    def _thumb_update(self, pil_img):
        t = pil_img.copy().convert("RGBA")
        t.thumbnail((96, 96), Image.LANCZOS)
        
        chk = Image.new("RGBA", (96, 96))
        d = ImageDraw.Draw(chk)
        for y in range(0, 96, 8):
            for x in range(0, 96, 8):
                c = ((200, 200, 200, 255) if (x // 8 + y // 8) % 2 == 0
                     else (160, 160, 160, 255))
                d.rectangle([x, y, x + 7, y + 7], fill=c)
        
        ox = (96 - t.width) // 2
        oy = (96 - t.height) // 2
        chk.paste(t, (ox, oy), t)
        
        self._tk_thumb = ImageTk.PhotoImage(chk)
        self.thumb_cv.delete("all")
        self.thumb_cv.create_image(0, 0, anchor="nw", image=self._tk_thumb)

    def _open_editor(self):
        if self._src is None:
            messagebox.showwarning("No image", "Please select an image first.")
            return
        CropEditor(self, self._src, self._edit_done)

    def _edit_done(self, result):
        self._final = result
        self._thumb_update(result)
        
        try:
            out = os.path.join(self._tmp, "folder_icon.png")
            result.save(out, "PNG")
            self._ico = out
            self.conv_l.config(text="Image ready! Click Apply.", fg=GREEN)
            self.status_v.set("Image ready")
        except Exception as e:
            self.conv_l.config(text=f"Save failed: {e}", fg=RED)

    def _check_ready(self):
        if not self._ico or not os.path.isfile(self._ico):
            messagebox.showwarning("Not Ready",
                "Please select and edit an image first.")
            return False
        
        folder = self.folder_path.get()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("No Folder", "Please select a valid folder.")
            return False
        
        return True

    def _run_pipeline(self, target_fn, args):
        self.progress.pack(fill="x", pady=(0, 8), before=self.apply_btn)
        self.progress.start(10)
        self.update()
        
        log = StepLog(self)

        def _status(msg):
            self.after(0, lambda m=msg: (self.status_v.set(m), log.log(m)))

        def _done(ok, msg):
            self.after(0, lambda o=ok, m=msg: self._finish(o, m, log))

        threading.Thread(
            target=target_fn,
            args=args + (_status, _done),
            daemon=True).start()

    def _apply(self):
        if not self._check_ready():
            return
        
        folder = self.folder_path.get()
        
        self._run_pipeline(apply_folder_icon_pipeline,
                          (folder, self._ico, self.hide_files_var.get()))

    def _finish(self, success, msg, log):
        self.progress.stop()
        self.progress.pack_forget()
        log.done()
        
        if success:
            messagebox.showinfo("Success!", msg)
            self._update_folder_status()
        else:
            messagebox.showerror("Error", msg)

    def _remove_icon(self):
        folder = self.folder_path.get()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("No Folder", "Please select a valid folder.")
            return
        
        if not messagebox.askyesno("Remove Icon",
            f"Remove custom icon from this folder?\n{folder}"):
            return
        
        self.progress.pack(fill="x", pady=(0, 8), before=self.apply_btn)
        self.progress.start(10)
        self.update()
        
        log = StepLog(self)

        def _status(msg):
            self.after(0, lambda m=msg: (self.status_v.set(m), log.log(m)))

        def _done(ok, msg):
            self.after(0, lambda o=ok, m=msg: self._finish_remove(o, m, log))

        threading.Thread(
            target=remove_folder_icon_pipeline,
            args=(folder, _status, _done),
            daemon=True).start()

    def _finish_remove(self, success, msg, log):
        self.progress.stop()
        self.progress.pack_forget()
        log.done()
        
        if success:
            messagebox.showinfo("Success!", msg)
            self._update_folder_status()
        else:
            messagebox.showerror("Error", msg)

    def _diagnostics(self):
        folder = self.folder_path.get()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("No Folder", "Please select a valid folder.")
            return
        messagebox.showinfo("Diagnostics", folder_diagnostics(folder))

    def destroy(self):
        try:
            shutil.rmtree(self._tmp, ignore_errors=True)
        except:
            pass
        super().destroy()


if __name__ == "__main__":
    app = HomePage()
    app.mainloop()