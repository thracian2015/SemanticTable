using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace SemanticTable
{
    internal sealed class ConnectedTableContext
    {
        public Excel.ListObject Table { get; set; }
        public Excel.QueryTable QueryTable { get; set; }
        public string ConnectionString { get; set; }
        public string MetadataConnectionString { get; set; }
        public string CommandText { get; set; }
    }

    internal static class ExcelConnectionService
    {
        public static ConnectedTableContext GetActiveSheetConnectedTable(Excel.Application app)
        {
            var cell = app.ActiveCell as Excel.Range;
            var sheet = app.ActiveSheet as Excel.Worksheet;
            if (cell == null || sheet == null)
                throw new InvalidOperationException("Open a worksheet and select a cell.");

            var selected = GetConnectedTableAtActiveCell(app);
            if (selected != null) return selected;
            var candidates = new System.Collections.Generic.List<ConnectedTableContext>();
            foreach (Excel.ListObject table in sheet.ListObjects)
            {
                var context = ContextFromTable(table);
                if (context != null) candidates.Add(context);
            }
            if (candidates.Count == 1) return candidates[0];
            if (candidates.Count > 1)
                throw new InvalidOperationException("This worksheet contains multiple semantic-model connected tables. Select a cell inside the table to use.");
            return null;
        }

        public static ConnectedTableContext GetConnectedTableAtActiveCell(Excel.Application app)
        {
            var cell = app.ActiveCell as Excel.Range;
            var sheet = app.ActiveSheet as Excel.Worksheet;
            if (cell == null || sheet == null) return null;
            foreach (Excel.ListObject table in sheet.ListObjects)
                try
                {
                    if (app.Intersect(cell, table.Range) == null) continue;
                    return ContextFromTable(table);
                }
                catch { }
            return null;
        }

        private static ConnectedTableContext ContextFromTable(Excel.ListObject table)
        {
            Excel.QueryTable queryTable;
            try { queryTable = table.QueryTable; }
            catch { return null; }
            if (queryTable == null) return null;
            string connection;
            try { connection = Convert.ToString(queryTable.WorkbookConnection.OLEDBConnection.Connection); }
            catch { connection = Convert.ToString(queryTable.Connection); }
            if (!IsSemanticModelConnection(connection)) return null;
            var command = ToCommandText(queryTable.CommandText);
            if (string.IsNullOrWhiteSpace(connection) || string.IsNullOrWhiteSpace(command)) return null;
            return new ConnectedTableContext
            {
                Table = table,
                QueryTable = queryTable,
                ConnectionString = connection,
                CommandText = command
            };
        }

        private static bool IsSemanticModelConnection(string connectionString)
        {
            var value = connectionString ?? string.Empty;
            return value.IndexOf("Provider=MSOLAP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Data Source=powerbi://", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Data Source=pbiazure://", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Data Source=pbidedicated://", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static ConnectedTableContext CreateConnectedTable(Excel.Application app, string connectionString)
        {
            var sheet = app.ActiveSheet as Excel.Worksheet;
            var destination = app.ActiveCell as Excel.Range;
            if (sheet == null || destination == null)
                throw new InvalidOperationException("Open an empty worksheet and select the destination cell.");
            if (!string.IsNullOrEmpty(Convert.ToString(destination.Value2)))
                throw new InvalidOperationException("Select an empty destination cell for the new connected table.");

            var excelConnection = connectionString.Trim();
            if (!excelConnection.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase))
                excelConnection = "OLEDB;" + excelConnection;
            const string initialDax = "EVALUATE ROW ( \"Semantic Table\", \"Select fields in Semantic Table Fields\" )";
            Excel.ListObject table = null;
            try
            {
                table = sheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcQuery, excelConnection,
                    Type.Missing, Excel.XlYesNoGuess.xlYes, destination);
                var queryTable = table.QueryTable;
                queryTable.BackgroundQuery = false;
                queryTable.CommandType = Excel.XlCmdType.xlCmdDefault;
                queryTable.CommandText = initialDax;
                if (!queryTable.Refresh(false)) throw new InvalidOperationException("Excel canceled creation of the connected table.");
                return new ConnectedTableContext
                {
                    Table = table,
                    QueryTable = queryTable,
                    ConnectionString = connectionString,
                    CommandText = initialDax
                };
            }
            catch
            {
                try { table?.Delete(); } catch { }
                throw;
            }
        }

        public static void ApplyAndRefresh(ConnectedTableContext context, string dax, int expectedColumnCount)
        {
            var operationId = Guid.NewGuid().ToString("N").Substring(0, 8);
            var queryTable = context.QueryTable;
            var workbookConnection = queryTable.WorkbookConnection;
            Excel.OLEDBConnection oleDb = null;
            object previousQueryCommand = null;
            object previousConnectionCommand = null;
            var wroteQueryTable = false;
            var wroteConnection = false;

            DiagnosticLog.Write("[" + operationId + "] Apply started. " +
                DescribeRefreshTarget(context, expectedColumnCount).Replace("\r\n", "; ") +
                "; Target: " + DescribeTarget(context.ConnectionString) +
                "; State: " + DescribeQueryState(queryTable) +
                "\r\n[" + operationId + "] DAX:\r\n" + dax);

            try { previousQueryCommand = queryTable.CommandText; }
            catch (Exception ex) { throw new InvalidOperationException("Could not read QueryTable.CommandText. " + ex.Message, ex); }

            try
            {
                oleDb = workbookConnection.OLEDBConnection;
                previousConnectionCommand = oleDb.CommandText;
            }
            catch { oleDb = null; }

            try
            {
                Exception queryWriteError = null;
                try
                {
                    queryTable.CommandText = CommandValueLike(previousQueryCommand, dax);
                    wroteQueryTable = true;
                    DiagnosticLog.Write("[" + operationId + "] Wrote QueryTable.CommandText. Previous command matched new command: " +
                        string.Equals(ToCommandText(previousQueryCommand), dax, StringComparison.Ordinal) + ".");
                }
                catch (Exception ex) { queryWriteError = ex; }

                if (!wroteQueryTable && oleDb != null)
                {
                    try
                    {
                        oleDb.CommandText = CommandValueLike(previousConnectionCommand, dax);
                        wroteConnection = true;
                        DiagnosticLog.Write("[" + operationId + "] Wrote OLEDBConnection.CommandText fallback.");
                    }
                    catch (Exception connectionWriteError)
                    {
                        throw new InvalidOperationException(
                            "Excel rejected the generated DAX while writing the query. QueryTable: " +
                            queryWriteError?.Message + " OLEDBConnection: " + connectionWriteError.Message, connectionWriteError);
                    }
                }

                if (!wroteQueryTable && !wroteConnection)
                    throw new InvalidOperationException(
                        "Excel rejected the generated DAX while writing QueryTable.CommandText. " + queryWriteError?.Message,
                        queryWriteError);

                bool refreshed;
                try
                {
                    DiagnosticLog.Write("[" + operationId + "] Calling QueryTable.Refresh(false). " + DescribeQueryState(queryTable));
                    refreshed = queryTable.Refresh(false);
                    DiagnosticLog.Write("[" + operationId + "] QueryTable.Refresh returned " + refreshed + ". " + DescribeQueryState(queryTable));
                }
                catch (Exception ex)
                {
                    DiagnosticLog.Write("[" + operationId + "] QueryTable.Refresh threw: " + ex);
                    var diagnostics = DescribeRefreshTarget(context, expectedColumnCount);
                    throw new InvalidOperationException(
                        "Excel accepted the generated DAX but failed while refreshing the connected table. " + ex.Message +
                        "\r\n\r\n" + diagnostics +
                        "\r\nDiagnostic log: " + DiagnosticLog.PathName +
                        "\r\n\r\nGenerated DAX:\r\n" + dax, ex);
                }
                if (!refreshed) throw new InvalidOperationException("Excel canceled the connected-table refresh.");
                context.CommandText = dax;
                DiagnosticLog.Write("[" + operationId + "] Apply completed successfully.");
            }
            catch
            {
                try
                {
                    if (wroteQueryTable) queryTable.CommandText = previousQueryCommand;
                    if (wroteConnection && oleDb != null) oleDb.CommandText = previousConnectionCommand;
                    DiagnosticLog.Write("[" + operationId + "] Restored the previous command after failure.");
                }
                catch (Exception rollbackError) { DiagnosticLog.Write("[" + operationId + "] Command rollback failed: " + rollbackError); }
                throw;
            }
        }

        private static string DescribeQueryState(Excel.QueryTable queryTable)
        {
            var parts = new System.Collections.Generic.List<string>();
            try { parts.Add("Connection=" + queryTable.WorkbookConnection.Name); } catch { }
            try { parts.Add("BackgroundQuery=" + queryTable.BackgroundQuery); } catch { }
            try { parts.Add("Refreshing=" + queryTable.Refreshing); } catch { }
            try { parts.Add("RefreshStyle=" + queryTable.RefreshStyle); } catch { }
            try { parts.Add("ResultRange=" + queryTable.ResultRange.Address[false, false]); } catch { }
            return string.Join(", ", parts);
        }

        private static string DescribeRefreshTarget(ConnectedTableContext context, int expectedColumnCount)
        {
            try
            {
                var table = context.Table;
                var range = table.Range;
                var currentColumns = table.ListColumns.Count;
                var details = "Excel table: " + table.Name + " (" + range.Address[false, false] + ")" +
                              "\r\nCurrent columns: " + currentColumns +
                              "\r\nRequested columns: " + expectedColumnCount;

                if (expectedColumnCount > currentColumns)
                {
                    var sheet = (Excel.Worksheet)table.Parent;
                    var firstColumn = range.Column + currentColumns;
                    var lastColumn = range.Column + expectedColumnCount - 1;
                    var firstRow = range.Row;
                    var lastRow = range.Row + range.Rows.Count - 1;
                    var expansion = sheet.Range[sheet.Cells[firstRow, firstColumn], sheet.Cells[lastRow, lastColumn]];
                    var nonEmpty = Convert.ToDouble(sheet.Application.WorksheetFunction.CountA(expansion));
                    details += "\r\nRequired expansion range: " + expansion.Address[false, false] +
                               "\r\nNon-empty cells in expansion range: " + nonEmpty;
                    if (nonEmpty > 0)
                        details += "\r\nClear or move the cells in that range, then try Apply again.";
                }
                else
                {
                    details += "\r\nThe refresh does not require additional worksheet columns. " +
                               "The failure may be caused by an Excel-incompatible result type or table constraint.";
                }
                return details;
            }
            catch (Exception diagnosticError)
            {
                return "Additional refresh diagnostics were unavailable: " + diagnosticError.Message;
            }
        }

        private static object CommandValueLike(object previous, string dax)
        {
            return previous is Array ? new[] { dax } : (object)dax;
        }

        public static string NormalizeForAdomd(string excelConnection)
        {
            var value = excelConnection ?? "";
            if (value.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase)) value = value.Substring(6);
            var safeParts = value.Split(';').Where(p =>
            {
                var part = p.TrimStart();
                return !part.StartsWith("Provider=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("Command Timeout=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("Integrated Security=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("Identity Provider=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("Password=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("PWD=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("User ID=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("UID=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("Access Token=", StringComparison.OrdinalIgnoreCase) &&
                       !part.StartsWith("Interactive Login=", StringComparison.OrdinalIgnoreCase);
            }).Where(p => !string.IsNullOrWhiteSpace(p));
            return string.Join(";", safeParts) + ";Interactive Login=Always";
        }

        public static string NormalizeForMsolap(string excelConnection)
        {
            var dataSource = GetProperty(excelConnection, "Data Source");
            var catalog = GetProperty(excelConnection, "Initial Catalog");
            if (string.IsNullOrWhiteSpace(dataSource))
                throw new InvalidOperationException("The Excel connection does not contain a Data Source.");

            var clean = "Provider=MSOLAP.8;Data Source=" + dataSource;
            if (!string.IsNullOrWhiteSpace(catalog)) clean += ";Initial Catalog=" + catalog;
            return clean + ";User ID=''";
        }

        public static string DescribeTarget(string connectionString)
        {
            var keys = new[] { "Data Source", "Location", "Initial Catalog", "Cube" };
            var parts = (connectionString ?? "").Split(';');
            return string.Join("; ", parts.Select(p => p.Trim()).Where(p =>
                keys.Any(k => p.StartsWith(k + "=", StringComparison.OrdinalIgnoreCase))));
        }

        public static bool UsesExcelPrivatePowerBiEndpoint(string connectionString) =>
            (connectionString ?? "").IndexOf("Data Source=pbiazure://api.powerbi.com", StringComparison.OrdinalIgnoreCase) >= 0 ||
            (connectionString ?? "").IndexOf("Data Source=pbidedicated://", StringComparison.OrdinalIgnoreCase) >= 0;

        public static string ReplaceDataSource(string connectionString, string workspaceEndpoint)
        {
            if (string.IsNullOrWhiteSpace(workspaceEndpoint) ||
                !workspaceEndpoint.Trim().StartsWith("powerbi://", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Enter a Power BI workspace connection beginning with powerbi://.");

            // DAX Studio may display '?readwrite' as a tool-specific connection
            // option. It is not part of the workspace server URL.
            workspaceEndpoint = workspaceEndpoint.Trim().Split('?')[0].TrimEnd('/');

            var parts = (connectionString ?? "").Split(';').Where(p =>
                !p.TrimStart().StartsWith("Location=", StringComparison.OrdinalIgnoreCase)).ToList();
            var replaced = false;
            for (var i = 0; i < parts.Count; i++)
                if (parts[i].TrimStart().StartsWith("Data Source=", StringComparison.OrdinalIgnoreCase))
                {
                    parts[i] = "Data Source=" + workspaceEndpoint;
                    replaced = true;
                }
            if (!replaced) parts.Add("Data Source=" + workspaceEndpoint);
            return string.Join(";", parts);
        }

        public static string GetProperty(string connectionString, string propertyName)
        {
            var prefix = propertyName + "=";
            var part = (connectionString ?? "").Split(';').Select(p => p.Trim())
                .FirstOrDefault(p => p.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
            return part == null ? null : part.Substring(prefix.Length).Trim().Trim('"', '\'');
        }

        public static string GetModelIdentity(string connectionString)
        {
            var catalog = GetProperty(connectionString, "Initial Catalog");
            if (string.IsNullOrWhiteSpace(catalog)) return null;
            const string prefix = "sobe_wowvirtualserver-";
            if (catalog.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                var candidate = catalog.Substring(prefix.Length);
                Guid id;
                if (Guid.TryParse(candidate, out id)) return id.ToString("D");
            }
            return catalog;
        }

        public static string SetProperty(string connectionString, string propertyName, string value)
        {
            var prefix = propertyName + "=";
            var parts = (connectionString ?? "").Split(';').Where(p =>
                !p.TrimStart().StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(p)).ToList();
            if (!string.IsNullOrWhiteSpace(value)) parts.Add(prefix + value);
            return string.Join(";", parts);
        }

        private static string ToCommandText(object value)
        {
            if (value is string text) return text;
            if (value is Array parts) return string.Join("", parts.Cast<object>().Select(Convert.ToString));
            return Convert.ToString(value);
        }
    }
}
