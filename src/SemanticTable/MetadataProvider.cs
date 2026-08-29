using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.AnalysisServices.AdomdClient;

namespace SemanticTable
{
    internal interface IMetadataProvider
    {
        IReadOnlyList<SemanticField> Load(ConnectedTableContext context);
    }

    // Uses Excel's already-authenticated MSOLAP/ADO session. This avoids opening a
    // second XMLA connection and therefore does not require the add-in to acquire,
    // store, or refresh a separate Microsoft Entra access token.
    internal sealed class ExcelAdoMetadataProvider : IMetadataProvider
    {
        public IReadOnlyList<SemanticHierarchy> Hierarchies { get; private set; } = Array.Empty<SemanticHierarchy>();

        public IReadOnlyList<SemanticField> Load(ConnectedTableContext context)
        {
            return LoadFromInteractiveMsolap(context);
        }

        private IReadOnlyList<SemanticField> LoadFromInteractiveMsolap(ConnectedTableContext context)
        {
            var result = new List<SemanticField>();
            try
            {
                var connectionString = ExcelConnectionService.NormalizeForMsolap(context.ConnectionString);
                var catalog = ExcelConnectionService.GetProperty(connectionString, "Initial Catalog");
                if (!string.IsNullOrWhiteSpace(catalog) && catalog.StartsWith("sobe_wowvirtualserver-", StringComparison.OrdinalIgnoreCase))
                {
                    var workbook = (Excel.Workbook)((Excel.Worksheet)context.Table.Parent).Parent;
                    var datasetId = ExtractDatasetId(catalog);
                    var savedCatalog = StateStore.LoadModelCatalog(workbook, datasetId);
                    connectionString = ExcelConnectionService.SetProperty(connectionString, "Initial Catalog", null);
                    var catalogs = DiscoverCatalogs(connectionString);
                    catalog = catalogs.FirstOrDefault(c => c.Values.Any(v =>
                        v.IndexOf(datasetId, StringComparison.OrdinalIgnoreCase) >= 0))?.Name;
                    if (catalog == null && CanOpenCatalog(connectionString, datasetId)) catalog = datasetId;
                    if (catalog == null)
                        catalog = catalogs.Select(c => c.Name).FirstOrDefault(c =>
                            string.Equals(c, savedCatalog, StringComparison.OrdinalIgnoreCase));
                    if (catalog == null) catalog = SelectCatalog(catalogs.Select(c => c.Name).ToList());
                    if (string.IsNullOrWhiteSpace(catalog)) throw new InvalidOperationException("Select a semantic model catalog.");
                    StateStore.SaveModelCatalog(workbook, datasetId, catalog);
                    connectionString = ExcelConnectionService.SetProperty(connectionString, "Initial Catalog", catalog);
                }

                using (var connection = new OleDbConnection(connectionString))
                {
                    context.MetadataConnectionString = connectionString;
                    connection.Open();
                    var tables = ReadOleDbTables(connection);
                    ReadOleDbFields(connection, tables, "$SYSTEM.TMSCHEMA_COLUMNS", SemanticFieldKind.Column, result);
                    ReadOleDbFields(connection, tables, "$SYSTEM.TMSCHEMA_MEASURES", SemanticFieldKind.Measure, result);
                    ResolveSortByColumns(result);
                    Hierarchies = ReadOleDbHierarchies(connection, tables, result);
                    result.RemoveAll(f => f.IsHidden);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "The native MSOLAP account-selection connection could not open the Power BI model. " + ex.Message +
                    "\r\n\r\nConnection target: " + ExcelConnectionService.DescribeTarget(context.ConnectionString), ex);
            }
            result.Sort((a, b) => string.Compare(a.Table + "\0" + a.Name, b.Table + "\0" + b.Name, StringComparison.CurrentCultureIgnoreCase));
            return result;
        }

        public IReadOnlyList<string> LoadDistinctValues(ConnectedTableContext context, SemanticField field, int limit = 500,
            bool descending = false, string search = null)
        {
            if (field == null || field.Kind != SemanticFieldKind.Column)
                throw new InvalidOperationException("Distinct values are available only for model columns.");
            if (string.IsNullOrWhiteSpace(context?.MetadataConnectionString))
                throw new InvalidOperationException("Open the Fields pane before loading filter values.");

            var column = "'" + field.Table.Replace("'", "''") + "'[" + field.Name.Replace("]", "]]") + "]";
            var sortColumn = string.IsNullOrWhiteSpace(field.SortByColumn) ? column :
                "'" + (field.SortByTable ?? field.Table).Replace("'", "''") + "'[" + field.SortByColumn.Replace("]", "]]" ) + "]";
            var direction = descending ? "DESC" : "ASC";
            var searchCondition = string.IsNullOrWhiteSpace(search) ? "" :
                " && CONTAINSSTRING(CONVERT(" + column + ", STRING), \"" + search.Replace("\"", "\"\"") + "\")";
            var source = string.Equals(column, sortColumn, StringComparison.OrdinalIgnoreCase)
                ? "FILTER(DISTINCT(" + column + "), NOT ISBLANK(" + column + ")" + searchCondition + ")"
                : "FILTER(SUMMARIZE('" + field.Table.Replace("'", "''") + "', " + column + ", " + sortColumn + "), NOT ISBLANK(" + column + ")" + searchCondition + ")";
            var dax = "EVALUATE TOPN(" + Math.Max(1, Math.Min(limit, 5000)) +
                      ", " + source + ", " + sortColumn + ", " + direction + ", " + column + ", ASC) " +
                      "ORDER BY " + sortColumn + " " + direction + ", " + column + " ASC";
            var values = new List<string>();
            using (var connection = new OleDbConnection(context.MetadataConnectionString))
            using (var command = new OleDbCommand(dax, connection))
            {
                connection.Open();
                using (var reader = command.ExecuteReader())
                    while (reader.Read())
                    {
                        var value = reader.GetValue(0);
                        values.Add(value is DateTime date ? date.ToString("yyyy-MM-dd") : Convert.ToString(value));
                    }
            }
            return values;
        }

        public void ValidateQuery(ConnectedTableContext context, string dax)
        {
            if (string.IsNullOrWhiteSpace(context?.MetadataConnectionString))
                throw new InvalidOperationException("The model validation connection is unavailable.");
            using (var connection = new OleDbConnection(context.MetadataConnectionString))
            using (var command = new OleDbCommand(dax, connection))
            {
                command.CommandTimeout = 120;
                connection.Open();
                using (var reader = command.ExecuteReader())
                    reader.Read();
            }
        }

        private sealed class CatalogInfo
        {
            public string Name { get; set; }
            public List<string> Values { get; } = new List<string>();
        }

        private static string ExtractDatasetId(string catalog)
        {
            var match = Regex.Match(catalog ?? "", @"sobe_wowvirtualserver-(?<id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
                throw new InvalidOperationException("The Excel connection's internal catalog does not contain a valid semantic-model identifier: " + catalog);
            return match.Groups["id"].Value;
        }

        private static bool CanOpenCatalog(string serverConnectionString, string catalog)
        {
            try
            {
                using (var connection = new OleDbConnection(
                    ExcelConnectionService.SetProperty(serverConnectionString, "Initial Catalog", catalog)))
                {
                    connection.Open();
                    return true;
                }
            }
            catch { return false; }
        }

        private static List<CatalogInfo> DiscoverCatalogs(string serverConnectionString)
        {
            using (var connection = new OleDbConnection(serverConnectionString))
            {
                connection.Open();
                var table = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Catalogs, null);
                if (table == null) return new List<CatalogInfo>();
                return table.Rows.Cast<DataRow>().Select(row =>
                    {
                        var info = new CatalogInfo { Name = Convert.ToString(row["CATALOG_NAME"]) };
                        foreach (DataColumn column in table.Columns)
                        {
                            var value = Convert.ToString(row[column]);
                            if (!string.IsNullOrWhiteSpace(value)) info.Values.Add(value);
                        }
                        return info;
                    })
                    .Where(c => !string.IsNullOrWhiteSpace(c.Name))
                    .GroupBy(c => c.Name, StringComparer.OrdinalIgnoreCase).Select(g => g.First())
                    .OrderBy(c => c.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
            }
        }

        private static string SelectCatalog(IReadOnlyList<string> catalogs)
        {
            if (catalogs.Count == 1) return catalogs[0];
            using (var dialog = new Form { Text = "Select Power BI semantic model", Width = 520, Height = 165,
                StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false, MaximizeBox = false })
            {
                var label = new Label { Left = 12, Top = 12, Width = 480, Text = "Select the semantic model used by this connected table:" };
                var combo = new ComboBox { Left = 12, Top = 38, Width = 480, DropDownStyle = ComboBoxStyle.DropDownList };
                combo.Items.AddRange(catalogs.Cast<object>().ToArray());
                if (combo.Items.Count > 0) combo.SelectedIndex = 0;
                var ok = new Button { Text = "OK", Left = 336, Top = 74, Width = 75, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "Cancel", Left = 417, Top = 74, Width = 75, DialogResult = DialogResult.Cancel };
                dialog.Controls.Add(label); dialog.Controls.Add(combo); dialog.Controls.Add(ok); dialog.Controls.Add(cancel);
                dialog.AcceptButton = ok; dialog.CancelButton = cancel;
                return dialog.ShowDialog() == DialogResult.OK ? Convert.ToString(combo.SelectedItem) : null;
            }
        }

        private static Dictionary<long, string> ReadOleDbTables(OleDbConnection connection)
        {
            var tables = new Dictionary<long, string>();
            using (var command = new OleDbCommand("SELECT * FROM $SYSTEM.TMSCHEMA_TABLES", connection))
            using (var reader = command.ExecuteReader())
                while (reader.Read())
                    if (!AsBool(OleDbValue(reader, "$SYSTEM.TMSCHEMA_TABLES", "IsHidden")))
                        tables[Convert.ToInt64(OleDbValue(reader, "$SYSTEM.TMSCHEMA_TABLES", "ID"))] =
                            Convert.ToString(OleDbValue(reader, "$SYSTEM.TMSCHEMA_TABLES", "Name", "TableName"));
            return tables;
        }

        private static void ReadOleDbFields(OleDbConnection connection, Dictionary<long, string> tables,
            string rowset, SemanticFieldKind kind, ICollection<SemanticField> result)
        {
            using (var command = new OleDbCommand($"SELECT * FROM {rowset}", connection))
            using (var reader = command.ExecuteReader())
                while (reader.Read())
                {
                    var tableId = Convert.ToInt64(OleDbValue(reader, rowset, "TableID"));
                    string table;
                    var hidden = AsBool(OleDbValue(reader, rowset, "IsHidden"));
                    if ((!hidden || kind == SemanticFieldKind.Column) && tables.TryGetValue(tableId, out table))
                        result.Add(new SemanticField
                        {
                            Table = table,
                            Name = Convert.ToString(OleDbValue(reader, rowset, "Name", kind == SemanticFieldKind.Column ? "ColumnName" : "MeasureName")),
                            DataType = Convert.ToString(TryOleDbValue(reader, "DataType", "Data_Type")),
                            DisplayFolder = Convert.ToString(TryOleDbValue(reader, "DisplayFolder", "Display_Folder")),
                            MetadataId = ToInt64(TryOleDbValue(reader, "ID")),
                            SortByColumnId = ToInt64(TryOleDbValue(reader, "SortByColumnID", "Sort_By_Column_ID")),
                            IsHidden = hidden,
                            Kind = kind
                        });
                }
        }

        private static long ToInt64(object value) => value == null || value == DBNull.Value ? 0 : Convert.ToInt64(value);

        private static void ResolveSortByColumns(IEnumerable<SemanticField> fields)
        {
            var columnsById = fields.Where(f => f.Kind == SemanticFieldKind.Column && f.MetadataId != 0)
                .GroupBy(f => f.MetadataId).ToDictionary(g => g.Key, g => g.First());
            foreach (var field in fields.Where(f => f.Kind == SemanticFieldKind.Column && f.SortByColumnId != 0))
            {
                SemanticField sort;
                if (!columnsById.TryGetValue(field.SortByColumnId, out sort)) continue;
                field.SortByTable = sort.Table;
                field.SortByColumn = sort.Name;
            }
        }

        private static IReadOnlyList<SemanticHierarchy> ReadOleDbHierarchies(OleDbConnection connection,
            Dictionary<long, string> tables, IEnumerable<SemanticField> fields)
        {
            var byId = fields.Where(f => f.Kind == SemanticFieldKind.Column && f.MetadataId != 0)
                .GroupBy(f => f.MetadataId).ToDictionary(g => g.Key, g => g.First());
            var hierarchies = new Dictionary<long, SemanticHierarchy>();
            using (var command = new OleDbCommand("SELECT * FROM $SYSTEM.TMSCHEMA_HIERARCHIES", connection))
            using (var reader = command.ExecuteReader())
                while (reader.Read())
                {
                    var tableId = ToInt64(TryOleDbValue(reader, "TableID"));
                    string table;
                    if (AsBool(TryOleDbValue(reader, "IsHidden")) || !tables.TryGetValue(tableId, out table)) continue;
                    var id = ToInt64(TryOleDbValue(reader, "ID"));
                    if (id == 0) continue;
                    hierarchies[id] = new SemanticHierarchy
                    {
                        Table = table,
                        Name = Convert.ToString(TryOleDbValue(reader, "Name")),
                        DisplayFolder = Convert.ToString(TryOleDbValue(reader, "DisplayFolder", "Display_Folder"))
                    };
                }

            var levels = new List<System.Tuple<long, int, long>>();
            using (var command = new OleDbCommand("SELECT * FROM $SYSTEM.TMSCHEMA_LEVELS", connection))
            using (var reader = command.ExecuteReader())
                while (reader.Read())
                {
                    if (AsBool(TryOleDbValue(reader, "IsHidden"))) continue;
                    var hierarchyId = ToInt64(TryOleDbValue(reader, "HierarchyID"));
                    var columnId = ToInt64(TryOleDbValue(reader, "ColumnID", "SourceColumnID"));
                    var ordinal = (int)ToInt64(TryOleDbValue(reader, "Ordinal"));
                    if (hierarchies.ContainsKey(hierarchyId) && byId.ContainsKey(columnId))
                        levels.Add(System.Tuple.Create(hierarchyId, ordinal, columnId));
                }
            foreach (var level in levels.OrderBy(l => l.Item2))
                hierarchies[level.Item1].Levels.Add(byId[level.Item3]);
            return hierarchies.Values.Where(h => h.Levels.Count > 0)
                .OrderBy(h => h.Table, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(h => h.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        private static object OleDbValue(OleDbDataReader reader, string rowset, params string[] requestedNames)
        {
            for (var i = 0; i < reader.FieldCount; i++)
                foreach (var requested in requestedNames)
                    if (string.Equals(reader.GetName(i), requested, StringComparison.OrdinalIgnoreCase))
                        return reader.GetValue(i);

            for (var i = 0; i < reader.FieldCount; i++)
            {
                var actual = NormalizeColumnName(reader.GetName(i));
                foreach (var requested in requestedNames)
                {
                    var wanted = NormalizeColumnName(requested);
                    if (actual == wanted || actual.EndsWith(wanted, StringComparison.OrdinalIgnoreCase))
                        return reader.GetValue(i);
                }
            }

            var returned = string.Join(", ", Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
            throw new InvalidOperationException(
                rowset + " did not return " + string.Join("/", requestedNames) + ". Returned columns: " + returned);
        }

        private static string NormalizeColumnName(string name)
        {
            if (string.IsNullOrEmpty(name)) return string.Empty;
            return new string(name.Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();
        }

        private static object TryOleDbValue(OleDbDataReader reader, params string[] requestedNames)
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                var actual = NormalizeColumnName(reader.GetName(i));
                foreach (var requested in requestedNames)
                {
                    var wanted = NormalizeColumnName(requested);
                    if (actual == wanted || actual.EndsWith(wanted, StringComparison.OrdinalIgnoreCase))
                        return reader.IsDBNull(i) ? null : reader.GetValue(i);
                }
            }
            return null;
        }

        private static IReadOnlyList<SemanticField> LoadFromInteractiveAdomd(ConnectedTableContext context)
        {
            var result = new List<SemanticField>();
            try
            {
                using (var connection = new AdomdConnection(ExcelConnectionService.NormalizeForAdomd(context.ConnectionString)))
                {
                    connection.Open();
                    var tables = ReadAdomdTables(connection);
                    ReadAdomdFields(connection, tables, "$SYSTEM.TMSCHEMA_COLUMNS", SemanticFieldKind.Column, result);
                    ReadAdomdFields(connection, tables, "$SYSTEM.TMSCHEMA_MEASURES", SemanticFieldKind.Measure, result);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Microsoft interactive sign-in could not open the Power BI XMLA connection. " + ex.Message, ex);
            }
            result.Sort((a, b) => string.Compare(a.Table + "\0" + a.Name, b.Table + "\0" + b.Name, StringComparison.CurrentCultureIgnoreCase));
            return result;
        }

        private static Dictionary<long, string> ReadAdomdTables(AdomdConnection connection)
        {
            var tables = new Dictionary<long, string>();
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "SELECT [ID], [Name], [IsHidden] FROM $SYSTEM.TMSCHEMA_TABLES";
                using (var reader = command.ExecuteReader())
                    while (reader.Read())
                        if (!AsBool(reader["IsHidden"]))
                            tables[Convert.ToInt64(reader["ID"])] = Convert.ToString(reader["Name"]);
            }
            return tables;
        }

        private static void ReadAdomdFields(AdomdConnection connection, Dictionary<long, string> tables,
            string rowset, SemanticFieldKind kind, ICollection<SemanticField> result)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = $"SELECT [TableID], [Name], [IsHidden], [DisplayFolder] FROM {rowset}";
                using (var reader = command.ExecuteReader())
                    while (reader.Read())
                    {
                        var tableId = Convert.ToInt64(reader["TableID"]);
                        string table;
                        if (!AsBool(reader["IsHidden"]) && tables.TryGetValue(tableId, out table))
                            result.Add(new SemanticField
                            {
                                Table = table,
                                Name = Convert.ToString(reader["Name"]),
                                DisplayFolder = Convert.ToString(reader["DisplayFolder"]),
                                Kind = kind
                            });
                    }
            }
        }

        private static IReadOnlyList<SemanticField> LoadFromExcelNativeFallbacks(ConnectedTableContext context)
        {
            Exception adoError;
            try { return LoadFromAdo(context); }
            catch (Exception ex) { adoError = ex; }

            Exception queryError;
            try { return LoadFromTemporaryQueryTable(context); }
            catch (Exception ex) { queryError = ex; }

            try { return LoadFromTemporaryPivotTable(context); }
            catch (Exception pivotError)
            {
                throw new InvalidOperationException(
                    "All Excel-native metadata methods failed. " +
                    "ADO: " + adoError.Message + " QueryTable: " + queryError.Message +
                    " PivotTable: " + pivotError.Message, pivotError);
            }
        }

        private static IReadOnlyList<SemanticField> LoadFromAdo(ConnectedTableContext context)
        {
            var result = new List<SemanticField>();
            Excel.OLEDBConnection oleDb;
            try { oleDb = context.QueryTable.WorkbookConnection.OLEDBConnection; }
            catch (Exception ex) { throw new InvalidOperationException("Could not access WorkbookConnection.OLEDBConnection: " + ex.Message, ex); }

            try { if (!oleDb.IsConnected) oleDb.MakeConnection(); }
            catch (Exception ex) { throw new InvalidOperationException("OLEDBConnection.MakeConnection failed: " + ex.Message, ex); }

            dynamic connection;
            try { connection = oleDb.ADOConnection; }
            catch (Exception ex) { throw new InvalidOperationException("OLEDBConnection.ADOConnection is unavailable: " + ex.Message, ex); }
            if (connection == null)
                throw new InvalidOperationException("Excel did not expose an active ADO connection. Refresh the connected table and try again.");

            var tables = ReadTables(connection);
            ReadFields(connection, tables, "$SYSTEM.TMSCHEMA_COLUMNS", SemanticFieldKind.Column, result);
            ReadFields(connection, tables, "$SYSTEM.TMSCHEMA_MEASURES", SemanticFieldKind.Measure, result);
            result.Sort((a, b) => string.Compare(a.Table + "\0" + a.Name, b.Table + "\0" + b.Name, StringComparison.CurrentCultureIgnoreCase));
            return result;
        }

        private static IReadOnlyList<SemanticField> LoadFromTemporaryQueryTable(ConnectedTableContext context)
        {
            var result = new List<SemanticField>();
            var workbook = (Excel.Workbook)((Excel.Worksheet)context.Table.Parent).Parent;
            var app = workbook.Application;
            Excel.Worksheet sheet = null;
            Excel.QueryTable query = null;
            var oldAlerts = app.DisplayAlerts;
            try
            {
                sheet = (Excel.Worksheet)workbook.Worksheets.Add();
                sheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                query = sheet.QueryTables.Add(context.QueryTable.WorkbookConnection, sheet.Range["A1"]);
                query.BackgroundQuery = false;
                query.RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells;
                query.CommandType = Excel.XlCmdType.xlCmdDefault;

                var tables = ReadTables(query);
                ReadFields(query, tables, "$SYSTEM.TMSCHEMA_COLUMNS", SemanticFieldKind.Column, result);
                ReadFields(query, tables, "$SYSTEM.TMSCHEMA_MEASURES", SemanticFieldKind.Measure, result);
                result.Sort((a, b) => string.Compare(a.Table + "\0" + a.Name, b.Table + "\0" + b.Name, StringComparison.CurrentCultureIgnoreCase));
                return result;
            }
            finally
            {
                try { query?.Delete(); } catch { }
                if (sheet != null)
                {
                    try { app.DisplayAlerts = false; sheet.Delete(); } catch { }
                    finally { app.DisplayAlerts = oldAlerts; }
                }
            }
        }

        private static IReadOnlyList<SemanticField> LoadFromTemporaryPivotTable(ConnectedTableContext context)
        {
            var result = new List<SemanticField>();
            var workbook = (Excel.Workbook)((Excel.Worksheet)context.Table.Parent).Parent;
            var app = workbook.Application;
            Excel.Worksheet sheet = null;
            Excel.PivotTable pivot = null;
            var oldAlerts = app.DisplayAlerts;
            try
            {
                sheet = (Excel.Worksheet)workbook.Worksheets.Add();
                sheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                var cache = workbook.PivotCaches().Create(
                    Excel.XlPivotTableSourceType.xlExternal,
                    context.QueryTable.WorkbookConnection,
                    Excel.XlPivotTableVersionList.xlPivotTableVersion15);
                pivot = cache.CreatePivotTable(sheet.Range["A1"], "_CTF_Metadata_" + DateTime.Now.Ticks);

                foreach (Excel.CubeField cubeField in pivot.CubeFields)
                {
                    if (!cubeField.ShowInFieldList) continue;
                    if (cubeField.CubeFieldType == Excel.XlCubeFieldType.xlMeasure)
                    {
                        result.Add(new SemanticField
                        {
                            Table = "Measures",
                            Name = cubeField.Caption,
                            Kind = SemanticFieldKind.Measure
                        });
                    }
                    else if (cubeField.CubeFieldType == Excel.XlCubeFieldType.xlHierarchy)
                    {
                        var tableName = FirstBracketedPart(cubeField.Name);
                        if (string.IsNullOrWhiteSpace(tableName)) tableName = "Fields";
                        result.Add(new SemanticField
                        {
                            Table = tableName,
                            Name = cubeField.Caption,
                            Kind = SemanticFieldKind.Column
                        });
                    }
                }

                result.Sort((a, b) => string.Compare(a.Table + "\0" + a.Name, b.Table + "\0" + b.Name, StringComparison.CurrentCultureIgnoreCase));
                return result;
            }
            finally
            {
                if (sheet != null)
                {
                    try { app.DisplayAlerts = false; sheet.Delete(); } catch { }
                    finally { app.DisplayAlerts = oldAlerts; }
                }
            }
        }

        private static string FirstBracketedPart(string uniqueName)
        {
            if (string.IsNullOrWhiteSpace(uniqueName) || uniqueName[0] != '[') return null;
            var end = uniqueName.IndexOf(']');
            return end > 1 ? uniqueName.Substring(1, end - 1).Replace("]]", "]") : null;
        }

        private static Dictionary<long, string> ReadTables(Excel.QueryTable query)
        {
            query.CommandText = "SELECT [ID], [Name], [IsHidden] FROM $SYSTEM.TMSCHEMA_TABLES";
            if (!query.Refresh(false)) throw new InvalidOperationException("Excel canceled the table-metadata query.");
            var rows = ReadRows(query.ResultRange);
            var tables = new Dictionary<long, string>();
            foreach (var row in rows)
                if (!AsBool(row["IsHidden"])) tables[Convert.ToInt64(row["ID"])] = Convert.ToString(row["Name"]);
            return tables;
        }

        private static void ReadFields(Excel.QueryTable query, Dictionary<long, string> tables,
            string rowset, SemanticFieldKind kind, ICollection<SemanticField> result)
        {
            query.CommandText = $"SELECT [TableID], [Name], [IsHidden], [DisplayFolder] FROM {rowset}";
            if (!query.Refresh(false)) throw new InvalidOperationException("Excel canceled the field-metadata query.");
            foreach (var row in ReadRows(query.ResultRange))
            {
                var tableId = Convert.ToInt64(row["TableID"]);
                string table;
                if (!AsBool(row["IsHidden"]) && tables.TryGetValue(tableId, out table))
                    result.Add(new SemanticField
                    {
                        Table = table,
                        Name = Convert.ToString(row["Name"]),
                        DisplayFolder = Convert.ToString(row["DisplayFolder"]),
                        Kind = kind
                    });
            }
        }

        private static IReadOnlyList<Dictionary<string, object>> ReadRows(Excel.Range range)
        {
            var result = new List<Dictionary<string, object>>();
            var values = range.Value2 as object[,];
            if (values == null || values.GetLength(0) < 2) return result;
            var rowStart = values.GetLowerBound(0);
            var rowEnd = values.GetUpperBound(0);
            var colStart = values.GetLowerBound(1);
            var colEnd = values.GetUpperBound(1);
            for (var r = rowStart + 1; r <= rowEnd; r++)
            {
                var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                for (var c = colStart; c <= colEnd; c++) row[Convert.ToString(values[rowStart, c])] = values[r, c];
                result.Add(row);
            }
            return result;
        }

        private static Dictionary<long, string> ReadTables(dynamic connection)
        {
            var tables = new Dictionary<long, string>();
            dynamic recordset = null;
            try
            {
                recordset = connection.Execute("SELECT [ID], [Name], [IsHidden] FROM $SYSTEM.TMSCHEMA_TABLES");
                while (!recordset.EOF)
                {
                    if (!AsBool(Field(recordset, "IsHidden")))
                        tables[Convert.ToInt64(Field(recordset, "ID"))] = Convert.ToString(Field(recordset, "Name"));
                    recordset.MoveNext();
                }
            }
            finally
            {
                CloseRecordset(recordset);
            }
            return tables;
        }

        private static void ReadFields(dynamic connection, Dictionary<long, string> tables,
            string rowset, SemanticFieldKind kind, ICollection<SemanticField> result)
        {
            dynamic recordset = null;
            try
            {
                recordset = connection.Execute($"SELECT [TableID], [Name], [IsHidden], [DisplayFolder] FROM {rowset}");
                while (!recordset.EOF)
                {
                    var tableId = Convert.ToInt64(Field(recordset, "TableID"));
                    var isHidden = AsBool(Field(recordset, "IsHidden"));
                    string table = null;
                    if (!isHidden && tables.TryGetValue(tableId, out table))
                        result.Add(new SemanticField
                        {
                            Table = table,
                            Name = Convert.ToString(Field(recordset, "Name")),
                            DisplayFolder = Convert.ToString(Field(recordset, "DisplayFolder")),
                            Kind = kind
                        });
                    recordset.MoveNext();
                }
            }
            finally { CloseRecordset(recordset); }
        }

        private static object Field(dynamic recordset, string name) => recordset.Fields.Item(name).Value;
        private static bool AsBool(object value) => value != null && value != DBNull.Value && Convert.ToBoolean(value);

        private static void CloseRecordset(dynamic recordset)
        {
            if (recordset == null) return;
            try { if (recordset.State != 0) recordset.Close(); } catch { }
            try { Marshal.FinalReleaseComObject(recordset); } catch { }
        }
    }
}
