# Third-Party Notices

Semantic Table depends on the components below. The MIT license in this repository covers Semantic Table’s source code only. It does not relicense third-party code or Microsoft components. Each dependency remains governed by its own license or distribution terms.

## Excel-DNA

- Package: `ExcelDna.AddIn` 1.8.0
- Copyright: Copyright (C) 2005-2023 Govert van Drimmelen
- License: zlib License
- Project: <https://excel-dna.net/>
- Upstream license: <https://github.com/Excel-DNA/ExcelDna/blob/master/Distribution/LICENSE.txt>

The upstream zlib notice must not be removed or altered from source distributions. Refer to the upstream license for the complete controlling text.

## Newtonsoft.Json

- Package: `Newtonsoft.Json` 13.0.3
- Copyright: Copyright (c) 2007 James Newton-King
- License: MIT License
- Project: <https://www.newtonsoft.com/json>
- Upstream license: <https://github.com/JamesNK/Newtonsoft.Json/blob/master/LICENSE.md>

The upstream copyright and permission notice must be included with copies or substantial portions of Newtonsoft.Json. Refer to the upstream license for the complete controlling text.

## Microsoft Office Interop

- Package: `Microsoft.Office.Interop.Excel` 15.0.4795.1001
- Publisher and copyright holder: Microsoft Corporation
- NuGet package: <https://www.nuget.org/packages/Microsoft.Office.Interop.Excel/15.0.4795.1001>

Microsoft Office Interop is a Microsoft component and is not covered by Semantic Table’s MIT license. The package does not grant a Microsoft Office license. Before publishing a packed XLL that embeds any Interop assembly, verify the applicable Microsoft package, Office, and redistribution terms for the intended distribution method.

Semantic Table uses the package as a compile-time interop definition and embeds the required interop types into its own assembly. The Microsoft Office Interop assembly is not intended to be included as a separate release binary or packed resource.

## Distribution review required

The source repository may reference these packages for development and local builds. Do not infer from a successful Excel-DNA pack operation that every embedded binary may be redistributed under MIT. Semantic Table requires the Microsoft Analysis Services OLE DB Provider (MSOLAP) to be installed separately and does not redistribute MSOLAP.

Microsoft, Microsoft Excel, Microsoft Office, Microsoft Fabric, and Power BI are trademarks of the Microsoft group of companies. Their use identifies interoperability and does not imply endorsement.
