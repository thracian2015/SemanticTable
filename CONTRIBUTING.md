# Contributing

Contributions are welcome after the public repository is available.

## Development setup

1. Use Windows with 64-bit desktop Excel.
2. Install Visual Studio 2022 with .NET desktop development tools and the .NET Framework 4.8 Developer Pack.
3. Restore and build:

   ```powershell
   dotnet restore SemanticTable.sln
   dotnet build SemanticTable.sln -c Release -p:Platform=x64 --no-restore
   ```

4. Test with a sanitized workbook and non-production semantic model.

## Pull requests

- Keep changes focused and preserve existing behavior unless the change is documented.
- Do not commit credentials, tokens, tenant/workspace/model identifiers, connection strings, local paths, or customer data.
- Use placeholders such as `<workspace>` and `<semantic-model>` in examples.
- Update `CHANGELOG.md` for user-visible changes.
- Confirm `Release | x64` builds with zero errors.
- Describe manual Excel testing, including the Excel and Microsoft 365 build used.

## Licensing

By submitting a contribution, you agree that it may be distributed under the repository’s MIT License. Do not submit third-party code unless its license and required notices are documented and compatible with the project’s distribution plan.
