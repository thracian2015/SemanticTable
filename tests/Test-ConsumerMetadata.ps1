$ErrorActionPreference = 'Stop'
$testSource = @'
using System;
using System.Data;
using System.Linq;
using SemanticTable;
public static class ConsumerMetadataTests
{
    private static DataTable Table(params string[] columns) {
        var t = new DataTable(); foreach (var c in columns) t.Columns.Add(c); return t;
    }
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    public static void Run() {
        var a = Table("DIMENSION_UNIQUE_NAME", "HIERARCHY_UNIQUE_NAME", "HIERARCHY_ORIGIN", "HIERARCHY_IS_VISIBLE", "HIERARCHY_DISPLAY_FOLDER");
        a.Rows.Add("[Da]]te]", "[Da]]te].[Year]", "2", "true", "Calendar");
        a.Rows.Add("[Da]]te]", "[Da]]te].[Month]", "2", "true", "Calendar");
        a.Rows.Add("[Da]]te]", "[Da]]te].[Private]", "2", "false", "");
        a.Rows.Add("[Da]]te]", "[Da]]te].[Calendar]", "1", "true", "Calendar");
        a.Rows.Add("[Da]]te]", "[Da]]te].[Unresolved]", "1", "true", "");
        var l = Table("HIERARCHY_UNIQUE_NAME", "LEVEL_TYPE", "LEVEL_NUMBER", "LEVEL_ATTRIBUTE_HIERARCHY_NAME", "LEVEL_DBTYPE", "LEVEL_IS_VISIBLE");
        l.Rows.Add("[Da]]te].[Year]", "1", "0", "", "130", "true");
        l.Rows.Add("[Da]]te].[Year]", "0", "1", "", "20", "true");
        l.Rows.Add("[Da]]te].[Month]", "0", "1", "", "130", "true");
        l.Rows.Add("[Da]]te].[Calendar]", "0", "2", "[Da]]te].[Month]", "130", "true");
        l.Rows.Add("[Da]]te].[Calendar]", "0", "1", "Year", "20", "true");
        l.Rows.Add("[Da]]te].[Unresolved]", "0", "1", "Missing", "130", "true");
        var m = Table("MEASURE_UNIQUE_NAME", "MEASURE_NAME", "MEASUREGROUP_NAME", "MEASURE_IS_VISIBLE", "MEASURE_DISPLAY_FOLDER", "DATA_TYPE", "MEASURE_AGGREGATOR");
        m.Rows.Add("[Measures].[Net ]] Amount]", "Translated caption", "Sales", "true", "Financial", "6", "127");
        m.Rows.Add("[Measures].[Hidden]", "Hidden", "Sales", "false", "", "5", "127");
        m.Rows.Add("[Measures].[Sum of Amount]", "Sum of Amount", "Sales", "true", "", "5", "1");
        var fields = ConsumerMetadata.Parse(a, l, m, out var hierarchies);
        Check(fields.Count == 3, "Unexpected field count or hidden field leak");
        Check(fields.Single(f => f.Name == "Year").Table == "Da]te", "Escaped table identifier");
        Check(fields.Single(f => f.Name == "Year").DataType == "Int64", "All level must not supply column type");
        Check(fields.Single(f => f.Kind == SemanticFieldKind.Measure).Name == "Net ] Amount", "Caption used as model identifier");
        Check(fields.Single(f => f.Kind == SemanticFieldKind.Measure).Table == "Sales", "Measure home table lost");
        Check(hierarchies.Count == 1 && hierarchies[0].Levels.Select(f => f.Name).SequenceEqual(new[] { "Year", "Month" }), "Hierarchy source mapping/order");
        Check(fields.All(f => f.SortByColumn == null), "Consumer discovery invented a sort column");
        Check(ConsumerMetadata.DbTypeName(11) == "Boolean" && ConsumerMetadata.DbTypeName(7) == "DateTime"
            && ConsumerMetadata.DbTypeName(5) == "Double" && ConsumerMetadata.DbTypeName(6) == "Decimal"
            && ConsumerMetadata.DbTypeName(130) == "String", "OLE DB types interpreted as tabular type enums");
        a.Rows[1]["HIERARCHY_IS_VISIBLE"] = "false";
        fields = ConsumerMetadata.Parse(a, l, m, out hierarchies);
        Check(hierarchies.Count == 0 && fields.Count == 2, "Partial hierarchy exposed with inaccessible source column");
        Check(ConsumerMetadata.IdentifierParts("[A.B].[C]]D]").SequenceEqual(new[] { "A.B", "C]D" }), "Escaped multipart name parsing");
    }
}
'@
# Compile separate files so their using directives remain at compilation-unit scope.
$tempFile = Join-Path ([IO.Path]::GetTempPath()) ('ConsumerMetadataTests-' + [Guid]::NewGuid() + '.cs')
try {
    Set-Content -LiteralPath $tempFile -Value $testSource
    Add-Type -Path "$PSScriptRoot\..\src\SemanticTable\Models.cs", "$PSScriptRoot\..\src\SemanticTable\ConsumerMetadata.cs", $tempFile -ReferencedAssemblies (Get-ChildItem (Join-Path $PSHOME 'ref\*.dll')).FullName
    [ConsumerMetadataTests]::Run()
    Write-Output 'PASS: consumer metadata identifiers, visibility, types, hierarchy mapping/order, missing source columns, and sort fallback.'
} finally { Remove-Item -LiteralPath $tempFile -ErrorAction SilentlyContinue }
