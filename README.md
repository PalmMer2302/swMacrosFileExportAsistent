# swMacrosFileExportAsistent

This project contains a SolidWorks VBA macro tool for automating the export of 2D drawing files.  
The macro is intended for use in internal documentation processes at RPT.

## 🧩 Project Overview

- **Macro Name:** 2D_Export - RPT.swp
- **Language:** VBA (Visual Basic for Applications)
- **Platform:** SolidWorks Macro Environment
- **Purpose:** Automate the export of 2D drawing views to file

## 📁 Project Structure

swMacrosFileExportAsistent/
├── ExportedMacros/ # Text-based VBA modules (.bas, .frm)
│ ├── Module1.bas
│ ├── SaveForm.frm
├── Macro/
│ └── 2D_Export - RPT.swp # Binary macro file for SolidWorks
├── .gitignore
├── README.md


## 🚀 Usage

1. Open SolidWorks
2. Open the macro: `Tools > Macro > Run > 2D_Export - RPT.swp`
3. Follow the UI prompts to export 2D views automatically

## 🛠 Development Notes

- Always export updated VBA modules (.bas / .frm) before committing
- The `.swp` file is not tracked in Git, but kept for execution purposes
- Use Git to manage and track changes in macro logic and UI forms

## 👤 Author

Patiphan Jampacome  
Senior Product Design Engineer  
RPTD Engineering Team

## 📜 License

For internal use only. Not licensed for external distribution.
