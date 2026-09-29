# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Contensive Patterns Reference

This project is built on the Contensive platform. Before implementing any Contensive-specific code (admin UI, addons, database models, collection XML, portals, page widgets, remote methods, etc.), you MUST read the relevant pattern documentation.

**Start here — read the patterns index to find the right pattern for your task:**
- [Contensive Patterns Index](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/index.md)

Then read the specific pattern(s) that apply to the work you are doing. Key patterns include:
- [AdminUI Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/adminui-pattern.md) — LayoutBuilderList, LayoutBuilderNameValue, filters, sorting, pagination, CSV export
- [Addon Collection Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/addon-collection-pattern.md) — collection XML, CDefs, Fields, packaging
- [Database Models Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/database-models-pattern.md) — typed C# model classes
- [Addon Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/addon-pattern.md) — addon architecture and types
- [Portal Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/portal-pattern.md) — portal-style admin interfaces
- [Page Widget Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/addon-page-widget-pattern.md) — design block widgets
- [Best Practices](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/best-practices-pattern.md) — error handling, try/catch, coding conventions
- [Security Best Practices](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/security-best-practices.md) — authentication, authorization, secure coding

Do NOT implement Contensive features from memory — always fetch and read the current pattern documentation first.

## GUIDs

When a GUID is needed (collection XML, addon definitions, CDefs, etc.), always generate a real unique GUID using the command line rather than hand-typing one:

```bash
python3 -c "import uuid; print(f'{{{uuid.uuid4()}}}'.upper())"
```

Never use patterned or sequential GUIDs like `{A1B2C3D4-...}` — they risk collisions.

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

**Target**: .NET Standard 2.0 (netstandard2.0), assembly is strong-named (addonman.snk).

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
- **`Constants.cs`** — Centralized GUIDs, button names, and CSS styles

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
