# Docx2PDF

<p align="center">
  <img src="Docx2PDFLive.UI/Resources/app_icon.png" alt="Docx2PDF Icon" width="128" height="128" />
  <br>
  <strong>A high-speed, real-time DOCX to PDF converter for Windows with 100% Word fidelity</strong>
  <br>
  <em>Created by Akib</em>
</p>

---

## Overview

**Docx2PDF** is a modern Windows desktop application that watches a selected folder and all its subfolders recursively. Whenever a `.docx` file is created or updated, it converts it into a `.pdf` file in real-time with 100% native Word fidelity.

---

## Features

- **Strict High-Fidelity 1-to-1 Conversion**: Powered by native Microsoft Word COM or LibreOffice headless. Retains 100% visual fidelity including table borders, paragraph rules, section divider lines, custom typography, and backgrounds (critical for resumes, contracts, and formal documents).
- **Ultra-Fast Performance**: Features a persistent warm engine architecture delivering conversions in **150ms–300ms** per document, running seamlessly even on lower-end hardware.
- **Zero Ghost Process Guarantee**: Enforces Windows Kernel Job Object protection (`JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`) and automated 90-second idle release to ensure zero lingering `WINWORD.EXE` or `soffice.exe` background processes.
- **Mandatory Prerequisite Gate**: Verifies at launch that either Microsoft Word or LibreOffice is installed. If neither is detected, a clean setup modal guides the user to install LibreOffice with a single click.
- **Standalone Single-File Architecture**: Self-extracting compressed single binary (`Docx2PDF.exe`) with zero external DLL requirements.
- **Zero-Touch Online Auto-Updater**: Automatically checks for updates and upgrades seamlessly with zero user effort. Offline launch is fully supported with a non-blocking graceful fallback so conversions can continue anytime.
- **Live Real-Time Monitoring**: Automatically detects added, modified, or renamed Word documents across nested subdirectories.
- **Background Tray Mode**: Minimizing or closing `Docx2PDF` hides it directly into the Windows System Tray, continuing conversions quietly in the background without cluttering your taskbar.
- **Single-Instance Protection**: Clicking the app shortcut while it is already running automatically unhides and brings the existing window to the foreground, preventing duplicate processes.
- **Instant Folder Switching**: Changing folders midway immediately cancels in-progress jobs and focuses resources on the new directory.
- **Low Footprint**: Bounded parallel workers keep memory footprint minimal.

---

## Getting Started

### Installation
1. Go to the [Releases](https://github.com/akibsaif39/Docx2PDF/releases) section.
2. Download **`Docx2PDF_Install.exe`**.
3. Run the installer and launch **Docx2PDF**.

### Requirements
- **Operating System**: Windows 10 or Windows 11 (64-bit)
- **Conversion Engine**: Microsoft Word (any version) or [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/) (free & open-source).

---

## License

This project is licensed under the [MIT License](LICENSE).
