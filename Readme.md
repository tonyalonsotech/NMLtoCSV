# NMLtoCSV
A lightweight app to convert Traktor playlists from NML to CSV format.

## Why would anyone even want to do this?
- Converting playlists to CSV can simplify the process of turning History Playlists into complete tracklists when sharing mixes.
- The CSV format is also easier to work with for other projects, such as generating reference tables for web applications.

## How does the app work?
- Use the **Browse** button to locate and load your NML playlist file, or drag and drop the file anywhere into the window.
- Leave the output file as the default, or click **Change** to choose a different folder and/or filename.
- Select which columns you want included in the CSV.
- Arrange the column order as needed.
- Click **Export CSV** to generate your file.

## Does the app maintain persistence?
- Yes. Your preferences are automatically saved and restored the next time you open the app.
- Preference storage currently uses the Windows AppData directory. Cross-platform support for macOS and Linux is planned.

<img width="889" height="732" alt="image" src="https://github.com/user-attachments/assets/0a4bd3ab-3799-4506-82dc-620b56ee47b2" />


## Version History
V1.0
- Initial Release

V1.1
- Added drag-and-drop support

v1.2
- Refined UI text to reduce redundancy

v1.3
- Decoupled column selection order from default column order

v1.4
- Replaced checkboxes with tile-based selection UI

v1.5 - Current Version
- Added automatic preference saving and loading
- Preferences are stored in the Windows %APPDATA% directory as nmltocsv_preferences.json
