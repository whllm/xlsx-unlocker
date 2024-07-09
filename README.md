# XLSX Unlocker

This powershell script iterates through individual sheets and removes protection.
Protection is only removed at the sheet level.
This does *nothing* for documents which are password protected or encrypted at the document level.


Usage:

Download the script and execute it by doing either:
    Right-click -> Run with Powershell
    Or running it directly from the terminal

```powershell
.\xlsx-unlocker.ps1 # Allows you to select from window dialog
```
```powershell
.\xlsx-unlocker -InputPath "yourfile.xlsx"
```

Outputs will be relative to the script, so it's recommended to place the script in its own folder. Your unlocked_myfile.xlsx will be placed in the same directory as the script.

Script by whllm. You are free to copy, modify, redistribute and/or use as you see fit. This is a free tool intended for internal use, and is offered without warranty or guarantee of any kind. 
