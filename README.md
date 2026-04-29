# Windows Explorer Hash Column

A native Rust Windows Shell Extension that adds a "File Hash (MD5)" column to Explorer's Details view.

## Features
- **Native & Fast**: Built with Rust and native COM interfaces.
- **Smart Hashing**: Skips files > max size to prevent slowdowns, default 10MB.
- **Targeted**: Only activates for specific files, defined by the user.

## Configuration
Customize behavior via Registry (`HKEY_CURRENT_USER\Software\WinExplorerHash`):
- **Extensions**: Create a String (`REG_SZ`) value `Extensions`. Set to a comma-separated list of extensions (e.g., `.pdf,.png,.jpg`). *(Note: you must re-run `build_and_register.ps1` or `regsvr32` to apply changes to this key)*.
- **Max File Size**: Create a DWORD/QWORD value `MaxFileSizeBytes`. Set to max size in bytes (e.g. `10485760` for 10MB).

## Installation
Requires Windows 10/11 and [Rust](https://rustup.rs/).

1. Run **PowerShell as Administrator** in this directory:
   ```powershell
   .\build_and_register.ps1
   ```
2. **Restart Windows Explorer** (via Task Manager).

### How to use
1. Open a folder in **Details** view.
2. Right-click any column header -> **More...**
3. Check **File Hash (MD5)** -> **OK**.

## Uninstallation
Run in Administrator PowerShell:
```powershell
regsvr32 /u .\target\release\explorer_hash.dll
```
Restart Windows Explorer, then you can delete the project folder.
