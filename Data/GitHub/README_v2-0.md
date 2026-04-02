# JBH Services Local Group Policy Tool

JBH Services Local Group Policy Tool is a PowerShell-based utility for applying a Local Group Policy baseline on Windows systems.

It applies required policy settings automatically, and optional policy settings are preconfigured to the current JBH default selections but can be reviewed, changed, previewed, and either included or skipped before changes are written. After the changes are applied, Local Group Policy is refreshed and the user can choose whether to restart immediately or restart later.

## Project Documents

- [Changelog](./Data/GitHub/CHANGELOG.md)
- [Roadmap](./Data/GitHub/ROADMAP.md)
- [License](LICENSE)
- [Requirements](#requirements)
- [How to use](#how-to-use)

## Features

- Applies the JBH Services Local Group Policy baseline
- Includes required policy settings that are always applied
- Includes optional policy settings that are preconfigured to the current JBH default selections, but can still be reviewed and changed before the tool applies them
- Provides policy preview and review screens before changes are written
- Automatically requests Administrator elevation when needed
- Displays progress while policies are being applied
- Lets the user choose whether to restart immediately or restart later

## Requirements

- Windows 10/11 Pro
- PowerShell 5.1
- Administrator rights

## Files

- `LocalGroupPolicyTool.ps1` — main PowerShell script
- `Launch-LocalGroupPolicyTool.vbs` — launcher for running the tool more easily

## What the tool does

This tool updates the local `Registry.pol` policy files for Computer Configuration and User Configuration, refreshes Local Group Policy, and applies the selected JBH Services baseline settings.

The tool is designed to make Local Group Policy changes easier to manage without manually searching through Group Policy Editor for every setting.

## How to use

1. Download or clone this repository.
2. Right-click and run the launcher or script.
3. Approve the Administrator prompt if shown.
4. Choose one of the available menu options:
   - use the JBH default policy selections
   - review or change optional policy selections
   - preview which policies will be changed
   - exit without making changes
5. Apply the selected settings.
6. Choose whether to restart immediately or restart later.

## Optional policy selections

The tool supports optional policy selections so the baseline can be adjusted before it is applied.

Optional items can be set to:

- **Apply** — include that policy in the current run
- **Skip** — leave that policy out of the current run

Optional settings follow the current JBH default selections unless the user changes them before applying.

## Notes

- Required policy settings are always applied.
- Optional policy settings follow the current JBH default selections unless the user changes them before applying.
- Running the tool again re-syncs the managed policy entries instead of stacking duplicates.
- Some settings may not fully take effect until Windows is restarted.

## Contributing

Suggestions, fixes, and improvements are welcome.

If you want to contribute:
1. Fork the repository
2. Make your changes
3. Open a pull request

All changes remain review-only until approved by the repository owner.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
