# JBH Services Local Group Policy Tool

JBH Services Local Group Policy Tool is a GUI-based Windows utility for applying a Local Group Policy baseline using a Python front end and a PowerShell backend.

The current version replaces the original console-based workflow with a click-based interface that lets you select policy settings visually, review them before applying, save and load configuration files, and then send the selected settings directly to the backend PowerShell script for processing.

## Project documents

- [Changelog](./Data/GitHub/CHANGELOG.md)
- [Roadmap](./Data/GitHub/ROADMAP.md)
- [Requirements](#requirements)
- [License](./Data/LICENSE)
- [Legacy Console Version README](./Data/GitHub/README_v2-0.md)

## Current version

This repository currently includes the newer GUI-based version of the tool.

### GUI version
- `LocalGroupPolicyTool_GUI.py`
- `LocalGroupPolicyTool.ps1`

### Legacy console version
- `Launch-LocalGroupPolicyTool_v2-0.vbs`
- `LocalGroupPolicyTool_v2-0.ps1`

The legacy console version is still included for compatibility and reference. See the [Legacy Console Version README](./Data/GitHub/README_v2-0.md) for information about the older workflow.

## Features

- GUI-based policy selection with checkboxes and dropdowns
- Review popup before applying settings
- Save and load configuration files
- Select All and Deselect All controls
- Automatic enable/disable handling for related settings
- PowerShell backend integration for applying selected policies
- Logging and backend result handling
- Restart recommendation after policy changes are applied

## How the current GUI version works

1. Open the GUI.
2. Select the policy settings you want to apply.
3. Adjust any available dropdown options, such as:
   - Automatic Updates option
   - Delivery Optimization mode
   - Feature Update deferral days
   - Scheduled install mode, day, and time
4. Click **Apply**.
5. Review the selected settings in the confirmation popup.
6. Click **OK** to continue.
7. The GUI sends the selected settings directly to the backend PowerShell script.
8. After the backend finishes, the GUI displays the result and lets you know if a restart is recommended.

## Configuration support

The GUI supports saving and loading configuration files so the same settings can be reused later or applied across multiple systems.

Saved configuration files are stored and loaded from the application's saved config location.

## Backend integration

The GUI does not directly write Local Group Policy changes itself.

Instead, it:
- collects the selected settings in the GUI
- prepares the configuration data
- calls the backend PowerShell script
- reads the backend result summary
- displays the final result to the user

This keeps the interface and the policy-writing logic separated.

## Logging

Backend runs create log data so apply results can be reviewed later if needed.

## Requirements

### For the GUI script version
- Windows 10/11 Pro
- Python 3
- PowerShell 5.1 or later
- Administrator rights when applying policy changes

### For packaged executable use
- Windows 10/11 Pro
- PowerShell 5.1 or later
- Administrator rights when applying policy changes

## Notes

- The GUI version is now the primary version of the tool.
- The legacy console version is still included for users who want the older workflow.
- Some policy changes may require a restart before they fully take effect.
- A restart recommendation is shown after the backend completes.

## Legacy console version

The original console version of the tool is still included in this repository.
