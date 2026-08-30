# Changelog

All notable changes to Semantic Table will be documented in this file.

The project follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Changed

- Moved binary distribution from the repository tree to downloadable GitHub Release assets.
- Updated installation and build documentation to link to Semantic Table's GitHub Releases page.
- Added a tag-triggered GitHub Actions workflow that builds the x64 add-in and publishes the consistently named XLL release asset.

## [0.1.0-beta.4] - 2026-08-30

### Changed

- Added the Semantic Table GitHub repository URL to the About window.

## [0.1.0-beta.3] - 2026-08-29

### Changed

- Standardized the distributable filename as `SemanticTable-win-x64.xll` so future releases can replace the existing file without changing Excel's add-in registration.
- Standardized the packed x64 XLL as a GitHub Release asset.

## [0.1.0-beta.2] - 2026-08-29

### Changed

- Renamed the project, solution, assembly, namespace, DNA manifest, and build outputs consistently to `SemanticTable`.
- Standardized hidden workbook-state identifiers on the `_SemanticTable_` prefix.
- Added repository, licensing, warranty, third-party notice, and roadmap documentation.
- Expanded the About window with version, copyright, MIT license, and warranty information.
- Expanded installation, authentication, licensing, build, screenshot, and known-limitation documentation.
- Added contributor guidance and a release-specific Microsoft binary redistribution warning.
- Removed the source-revision suffix from the About version and updated the copyright holder to Prologika, LLC.
- Removed the unused ADOMD.NET implementation and NuGet dependency; documented the separately installed 64-bit MSOLAP runtime prerequisite.
- Embedded required Office interop type metadata, packed Newtonsoft.Json, and produced a single-file unsigned x64 community artifact without Microsoft runtime DLLs.

## [0.1.0-beta.1] - 2026-08-29

### Added

- Initial public beta for building and filtering Excel connected tables backed by Power BI semantic models.
- Semantic-model metadata discovery, DAX generation, table refresh, filters, row limits, settings, and persisted workbook state.
