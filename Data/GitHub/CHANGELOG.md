# Changelog

All notable changes to this project will be documented in this file.

---
#### [2.5.1]
- Added the ability to save and load configuration files.
- Added the updated application folder structure.
- Added GUI-to-PowerShell backend integration so the GUI can send selected settings directly to the backend script.
- Added backend result handling in the GUI.
- Added logging path support for backend runs.
- Added detection and handling for the backend PowerShell script path.
- Added apply review and backend result popups in the GUI.
- Added must run the .exe as admin (right click "Run as admin").

#### [2.5.0]
- Added the first GUI version of the tool.
- Replaced console-based setting selection with a click-based interface.
- Added checkboxes and combo boxes for policy selection and configuration.
- Added dependency handling for related settings, such as Delivery Optimization, Feature Update deferral, and Automatic Updates scheduling.
- Added Select All and Deselect All controls.
- Added a visual layout for policy sections across multiple columns.

#### [2.0.0] 
- Initial public release of the JBH Services Local Group Policy Tool.

#### [1.6.0] 
- Final pre-public internal release used to consolidate features, improve usability, and prepare the tool for public release.

#### [1.5.4] 
- Standardized user-facing text throughout the tool for consistency and better guidance.
#### [1.5.3] 
- Made additional prompt and menu wording improvements for clarity and ease of use.
#### [1.5.2] 
- Further refined the user interface and screen flow for clearer navigation and review.
#### [1.5.1] 
- Improved handling for possible file access and permission issues when working with registry.pol files.
#### [1.5.0] 
- Expanded pre-apply review and exit options so the user has more control before changes are written.

#### [1.4.3] 
- Improved the progress display to show more detailed information about each policy being applied.
#### [1.4.2] 
- Added a progress indicator during policy application.
#### [1.4.1] 
- Added support for choosing whether to restart immediately after applying changes or restart later.
#### [1.4.0] 
- Added a confirmation step before changes are applied to reduce the risk of accidental policy changes.

#### [1.3.5] 
- Improved prompt wording and error messaging to provide clearer guidance during use.
#### [1.3.4] 
- Added more internal comments to better explain the purpose of each section and function.
#### [1.3.3] 
- Updated the registry.pol read/write handling to be more robust and reliable.
#### [1.3.2] 
- Improved the user interface and selection flow for clearer prompts, better screen organization, and easier review before applying changes.
#### [1.3.1] 
- Improved variable and function naming and added more descriptive comments for clarity.
#### [1.3.0] 
- Added comments and organized the script into sections to improve readability and maintenance.

#### [1.2.2] 
- Set a fixed console window size and locked resizing to keep the display layout consistent.
#### [1.2.1] 
- Added error handling to stop execution and display an error message if a failure occurs during processing.
#### [1.2.0] 
- Added an administrative privilege check and automatic elevation restart when the script is not launched as Administrator.

#### [1.1.5] 
- Clarified prompts and menu wording so that Skip clearly means the optional policy will not be included in the current run.
#### [1.1.4] 
- Added a summary screen after optional policy configuration so the user can review selections one final time before applying changes.
#### [1.1.3] 
- Updated the optional policy configuration flow to allow returning to the Main Menu without saving in-progress configuration changes.
#### [1.1.2] 
- Expanded the Main Menu descriptions and improved the optional policy configuration screen to show both the JBH default policy state and the current selection state.
#### [1.1.1] 
- Improved the optional policy selection flow to make reviewing and changing selections easier before applying changes.
#### [1.1.0] 
- Added a policy preview screen so the user can review which policies will be changed before confirming.

#### [1.0.2] 
- Updated optional policy selections to better match the current JBH baseline and removed policies no longer included in the baseline.
#### [1.0.1] 
- Added optional policy selections for Sleep option, TaskView button, and Chat icon.
#### [1.0.0] 
- Initial internal release of the JBH Services Local Group Policy Tool.
