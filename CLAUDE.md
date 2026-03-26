# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Architecture References
**On every session start, read the following before any task:**
- [Contensive Architecture](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/contensive-architecture.md)
- [Patterns Index](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/index.md)

## Architecture Overview
This project uses the Contensive framework. Key conventions:
- Use string interpolation, not concatenation  

## Project Overview

This is the **Add-on Manager** for the Contensive CMS platform. It's a C# addon (plugin) that provides a UI for administrators to install, export, upload, and uninstall addon collections within a Contensive site.

## Build Commands

```bash
# Build the project
dotnet build source/addonManager51/addonManager51.csproj --configuration Debug

# Clean
dotnet clean source/addonManager51/addonManager51.csproj

# Full build + package (uses 7-Zip for collection packaging)
scripts/build.cmd
```

There are no tests or linting configured in this project.

## Architecture

**Target**: .NET Framework 4.8 (net48), assembly is strong-named (addonman.snk).

**Key dependencies**: `Contensive.CPBaseClass` and `Contensive.DbModels` (NuGet).

### Source Layout (`source/addonManager51/`)

- **Addons/** — Classes extending `AddonBaseClass` that serve as entry points. Each addon has an `Execute(CPBaseClass cp)` method called by the Contensive framework:
  - `CollectionLibraryClass` — Main UI listing installed/available collections with install/upgrade/remove actions
  - `UploadClass` — Handles file upload and installation of collections
  - `ExportClass` — Packages and exports a collection as a downloadable zip
  - `UninstallClass` — Removes collections and cleans up DB records/navigator entries
  - `OnInstallClass` — Lifecycle hook that runs when this addon itself is installed
- **Controllers/** — Business logic separated from UI:
  - `InstallController` — Orchestrates collection installation from library, folder, or upload
  - `GenericController` — DB operations, navigator entry management, utility methods
- **`_Constants.cs`** — Centralized GUIDs, button names, and CSS styles

### UI (`ui/`)

- `AddonManagerLibraryBody.html` — Mustache-templated HTML for the collection library cards. Referenced by `CollectionLibraryClass` via a layout GUID.

### Collections (`collections/Add-on Manager/`)

- `Add-on Manager.xml` — Collection manifest defining addons, navigator entries, and metadata
- The build script packages the DLL + UI into zips here for deployment

## Contensive Patterns

- **Admin UI**: Built using `cp.AdminUI.CreateLayoutBuilder()` for forms and `cp.Html` helpers for HTML generation
- **Content records**: CRUD via `cp.Content.*` and `cp.CSNew()` (content set) methods
- **Error handling**: Exceptions reported via `cp.Site.ErrorReport(ex)`
- **Navigator entries**: Managed through content records in "Navigator Entries" table
