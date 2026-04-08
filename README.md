# Smart_BIMs

A C# Class Library plugin targeted for Revit 2024.

## Features
- **Multi-Version Support**: Works seamlessly on Revit 2024 (.NET 4.8) and Revit 2025/2026/2027 (.NET 8.0) natively!
- **Easy Schedule**: Quickly auto-generate schedules for multiple categories via a simple WPF UI.

## Local Setup
1. Open `Smart_BIMs.csproj` in Visual Studio.
2. Build the project.
3. Create a shortcut or copy `Smart_BIMs.addin` and the built `Smart_BIMs.dll` to your `%appdata%\Autodesk\Revit\Addins\2024` folder.

## Distribution & Auto-Update Strategy
This plugin utilizes a GitHub release checker in `App.cs`. When users launch Revit, it will check the `releases/latest` API on GitHub. If a newer version is present, it prompts the user to download the latest installer. So you just need to:
1. Push this to a public GitHub repo.
2. Publish MSI/Setup installers in the **Releases** tab.
3. Users get notified in Revit automatically when you push a new release!
