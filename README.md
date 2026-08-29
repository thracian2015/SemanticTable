# Semantic Table

Semantic Table 0.1.0-beta.1 is a Windows Excel add-in that gives Power BI connected query tables a PivotTable-like field picker. It reads semantic-model metadata, generates DAX, updates an Excel `QueryTable`, refreshes the table, and saves the selected fields and filters in the workbook.

This is beta software. Test it with non-production workbooks and semantic models before broader deployment.

## What it does

- Opens a field pane for a Power BI semantic-model connected table.
- Discovers visible model tables, columns, measures, and hierarchies.
- Builds a bounded DAX query from selected fields and filters.
- Refreshes the existing Excel table without replacing its `ListObject`.
- Restores the previous DAX command if refresh throws an error.
- Stores table settings in hidden workbook-level names beginning with `_SemanticTable_`.

## Screenshots

Product screenshots have not yet been captured for the public beta. The repository reserves [`docs/screenshots`](docs/screenshots/README.md) for sanitized captures of the ribbon, Fields pane, connection dialog, and filtered output. Do not add screenshots containing tenant names, workspace names, semantic-model names, or workbook data.

## Requirements

- Windows desktop Excel, 64-bit.
- .NET Framework 4.8.
- An existing Excel table connected to a Power BI semantic model, or a valid MSOLAP connection using placeholder-free values.
- The Power BI tenant setting that permits live Excel connections.
- Build permission on the semantic model, or at least Contributor access to its workspace.
- A Power BI/Fabric license and capacity configuration that permits the user to access the semantic model from Excel.

## Installation

There is no signed installer in the beta repository.

1. Obtain or build `SemanticTable64-packed.xll`.
2. In Excel, open **File > Options > Add-ins**.
3. At the bottom, select **Excel Add-ins**, choose **Go**, and then **Browse**.
4. Select `SemanticTable64-packed.xll`.
5. If Windows or organizational policy blocks the file, use only an approved trusted location or deployment method. Do not bypass endpoint-security policy.

The add-in appears on the **Semantic Table** ribbon tab. Select a cell in a supported connected table and choose **Fields**.

## Connections and authentication

Semantic Table uses the live, authenticated MSOLAP/ADO connection associated with the Excel workbook. It does not implement a separate interactive sign-in flow and does not request or persist a separate Microsoft Entra access token.

For a new connection, the dialog starts with placeholders only:

```text
Provider=MSOLAP.8;Data Source=powerbi://api.powerbi.com/v1.0/myorg/<workspace>;Initial Catalog=<semantic-model>
```

Replace both placeholders before connecting. The complete connection string is saved in a hidden workbook-level name for that table. Do not place passwords, access tokens, or other secrets in the connection string. Workbook recipients still need their own permitted Power BI identity; normal semantic-model permissions and row-level security continue to apply.

## Power BI and Fabric licensing limitations

Semantic Table does not grant a Power BI, Microsoft Fabric, or Microsoft 365 license and cannot override tenant, workspace, semantic-model, capacity, or row-level security settings.

Microsoft documents these baseline requirements for Excel live connections:

- The tenant administrator must allow users to work with Power BI semantic models in Excel using a live connection.
- The user needs semantic-model Build permission or at least Contributor access to the workspace.
- The user needs an applicable Power BI license.
- Fabric Free users are limited to semantic models in **My workspace** or qualifying Premium/Fabric capacity; Microsoft currently documents Fabric F64 or greater for this scenario.
- PPU workspace content requires a PPU license, including XMLA and Analyze in Excel access.
- Composite models can require Build permission or Contributor rights on upstream semantic models when accessed through XMLA.

Licensing and service behavior can change. Confirm current requirements in Microsoft’s [Excel live-connection prerequisites](https://learn.microsoft.com/power-bi/collaborate-share/service-connect-power-bi-datasets-excel) and [Fabric licensing documentation](https://learn.microsoft.com/fabric/enterprise/licenses) before deployment.

## Build

### Prerequisites

- Visual Studio 2022 with .NET desktop development tools, or a compatible MSBuild/.NET SDK installation.
- .NET Framework 4.8 Developer Pack.
- Network access to restore the packages listed in `src/SemanticTable/SemanticTable.csproj`, unless they are already cached.

### Command line

```powershell
dotnet restore SemanticTable.sln
dotnet build SemanticTable.sln -c Release -p:Platform=x64 --no-restore
```

The packed 64-bit output is:

```text
src\SemanticTable\bin\x64\Release\net48\publish\SemanticTable64-packed.xll
```

You can also open `SemanticTable.sln` in Visual Studio and build `Release | x64`.

## Known limitations

- The add-in targets 64-bit Windows desktop Excel and .NET Framework 4.8.
- Existing connected tables are the best-tested path; connection layouts can differ across Microsoft 365 builds.
- Selected-field order follows semantic-model table and field order; drag-and-drop ordering is not implemented.
- Columns become grouping keys and measures are evaluated at that grouping. The output is a flat summary, not guaranteed transaction-level detail.
- Column filters support typed dates, equality, inequality, text contains, ranges, and drag-and-drop placement. Measure filters and sort expressions are not implemented.
- Calculation-group handling and explicit grain configuration are incomplete.
- Large queries remain subject to Excel, Power BI/Fabric capacity, XMLA, timeout, and model limits.
- The XLL is not currently code-signed or distributed through an installer.
- The exact redistribution requirements for Microsoft binaries included in a packed XLL must be reviewed before publishing binaries.

## License, warranty, and dependencies

Semantic Table source code is licensed under the [MIT License](LICENSE). The software is provided **as-is, without warranty of any kind, express or implied**.

Third-party components are not relicensed under Semantic Table’s MIT license. See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md), especially the Microsoft redistribution review required before publishing a packed binary.

## Project documentation

- [Changelog](CHANGELOG.md)
- [Roadmap](ROADMAP.md)
- [Contributing](CONTRIBUTING.md)
- [Third-party notices](THIRD-PARTY-NOTICES.md)
