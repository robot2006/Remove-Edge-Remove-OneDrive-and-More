# Windows Utility Toolkit
![Windows Utility Toolkit Screenshot](./assets/toolkit-screenshot.png)

A powerful, all-in-one PowerShell GUI toolkit for cleaning and managing Windows 10/11 systems. Easily remove bloatware like Edge and OneDrive, reset the icon cache, find/kill processes locking a folder, and toggle an annoying Defender privacy setting/notification.

---

## Features

This toolkit combines several powerful scripts into a single, user-friendly interface with three main sections:

#### 1. System Cleanup
- **Permanently Remove Microsoft Edge:** Forcefully uninstalls the Edge browser, its services, and scheduled tasks.
- **Permanently Remove OneDrive:** Uninstalls the OneDrive client, removes it from the File Explorer sidebar, and blocks reinstallation.

#### 2. System Tools
- **Reset Icon Cache:** A quick tool to fix corrupted or broken desktop and folder icons by clearing the cache and restarting Explorer.
- **Toggle Defender Privacy:** A reversible tweak to enable or disable "Automatic sample submission" in Windows Defender.

#### 3. Folder Unlocker
- **Find & Kill Locking Processes:** Select a folder to identify which processes are currently using files within it.
- **Force-Close Handles:** Forcefully terminate all identified processes to unlock the folder, allowing it to be moved or deleted.

---

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- **Administrator Privileges** (the included launcher will handle this)
- **`handle64.exe`** from Microsoft Sysinternals for the "Folder Unlocker" tab to function.

---

## Tested On

- **OS:** Microsoft Windows 11 Pro
- **Version:** 24H2 (OS Build 10.0.26100)

---

## How to Use

1.  **Download:** Clone this repository or download the ZIP file and extract it.
2.  **Prerequisite:** Download `handle64.exe` and paste it into the `Force_Close_Folder` directory. You can download it from here:
	(https://download.sysinternals.com/files/Handle.zip).
			
3.  **Run the Toolkit:** Double-click the **`Run.bat`** file. It will automatically request administrator privileges and launch the GUI.

---

## ⚠️ WARNING ⚠️

This toolkit performs **destructive and irreversible actions**. Removing system components like Microsoft Edge or force-closing processes can lead to data loss or system instability.

**Use these tools at your own risk.** The author assumes no liability for data loss or system damage resulting from the use of this software. Always back up important data before running this script.
