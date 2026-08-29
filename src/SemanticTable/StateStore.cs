using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SemanticTable
{
    internal static class StateStore
    {
        private const string Prefix = "_SemanticTable_";
        private const string ConnectionName = "_SemanticTable_Connection_";
        private const string CatalogName = "_SemanticTable_XmlaCatalog_";

        private sealed class CompactDefinition
        {
            [JsonProperty("v")] public int Version { get; set; } = 6;
            [JsonProperty("d")] public string DatasetId { get; set; }
            [JsonProperty("t")] public string ExcelTableName { get; set; }
            [JsonProperty("f")] public List<string> SelectedFieldKeys { get; set; } = new List<string>();
            [JsonProperty("x")] public List<CompactFilter> Filters { get; set; } = new List<CompactFilter>();
            [JsonProperty("r")] public int RowLimit { get; set; } = 500000;
            [JsonProperty("u")] public bool DeferUpdate { get; set; } = true;
        }

        private sealed class CompactFilter
        {
            [JsonProperty("k")] public string FieldKey { get; set; }
            [JsonProperty("m")] public string Mode { get; set; }
            [JsonProperty("o")] public string Operator { get; set; }
            [JsonProperty("a")] public string Value { get; set; }
            [JsonProperty("b")] public string Value2 { get; set; }
            [JsonProperty("s")] public List<string> Values { get; set; } = new List<string>();
        }

        public static TableDefinition Load(Excel.Workbook workbook, string tableName)
        {
            try
            {
                var nameText = Prefix + tableName;
                var stored = LoadStringName(workbook, nameText);
                if (string.IsNullOrEmpty(stored)) throw new InvalidOperationException("No saved table definition exists.");

                // Read compatibility for the short-lived 1.9.0 chunked format.
                if (stored.StartsWith("b64:", StringComparison.OrdinalIgnoreCase))
                {
                    int count;
                    if (!int.TryParse(stored.Substring(4), out count) || count < 1 || count > 1000)
                        throw new InvalidOperationException("The saved table definition has an invalid chunk count.");
                    var base64 = new StringBuilder();
                    for (var index = 1; index <= count; index++)
                    {
                        var chunk = LoadStringName(workbook, StatePartName(tableName, index));
                        if (chunk == null) throw new InvalidOperationException("A saved table-definition chunk is missing.");
                        base64.Append(chunk);
                    }
                    stored = Encoding.UTF8.GetString(Convert.FromBase64String(base64.ToString()));
                }

                var root = JObject.Parse(stored);
                if (root.Property("v") == null)
                    return JsonConvert.DeserializeObject<TableDefinition>(stored);

                var compact = root.ToObject<CompactDefinition>();
                return new TableDefinition
                {
                    Version = compact.Version,
                    DatasetId = compact.DatasetId,
                    ExcelTableName = compact.ExcelTableName ?? tableName,
                    Fields = (compact.SelectedFieldKeys ?? new List<string>()).Select(FieldFromKey).Where(f => f != null).ToList(),
                    Filters = (compact.Filters ?? new List<CompactFilter>()).Select(f => new FieldFilter
                    {
                        Field = FieldFromKey(f.FieldKey),
                        Mode = f.Mode,
                        Operator = f.Operator,
                        Value = f.Value,
                        Value2 = f.Value2,
                        Values = f.Values ?? new List<string>()
                    }).Where(f => f.Field != null).ToList(),
                    RowLimit = compact.RowLimit,
                    DeferUpdate = compact.DeferUpdate
                };
            }
            catch { return new TableDefinition { ExcelTableName = tableName }; }
        }

        public static void Save(Excel.Workbook workbook, TableDefinition definition)
        {
            var nameText = Prefix + definition.ExcelTableName;
            var previousCount = ParseChunkCount(LoadStringName(workbook, nameText));
            var compact = new CompactDefinition
            {
                DatasetId = definition.DatasetId,
                ExcelTableName = definition.ExcelTableName,
                SelectedFieldKeys = (definition.Fields ?? new List<SemanticField>())
                    .Where(f => f != null).Select(f => f.Key).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                Filters = (definition.Filters ?? new List<FieldFilter>()).Where(f => f?.Field != null)
                    .Select(f => new CompactFilter
                    {
                        FieldKey = f.Field.Key,
                        Mode = f.Mode,
                        Operator = f.Operator,
                        Value = f.Value,
                        Value2 = f.Value2,
                        Values = f.Values ?? new List<string>()
                    }).ToList(),
                RowLimit = definition.RowLimit,
                DeferUpdate = definition.DeferUpdate
            };

            SaveStringName(workbook, nameText, JsonConvert.SerializeObject(compact, Formatting.None));
            for (var index = 1; index <= previousCount; index++)
                try { workbook.Names.Item(StatePartName(definition.ExcelTableName, index)).Delete(); } catch { }
        }

        public static string LoadConnectionString(Excel.Workbook workbook, string tableName) =>
            LoadStringName(workbook, ConnectionName + tableName);
        public static void SaveConnectionString(Excel.Workbook workbook, string tableName, string connectionString) =>
            SaveStringName(workbook, ConnectionName + tableName, connectionString);
        public static string LoadModelCatalog(Excel.Workbook workbook, string datasetId) =>
            LoadStringName(workbook, CatalogName + SafeNameToken(datasetId));
        public static void SaveModelCatalog(Excel.Workbook workbook, string datasetId, string catalog) =>
            SaveStringName(workbook, CatalogName + SafeNameToken(datasetId), catalog);

        private static SemanticField FieldFromKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return null;
            var first = key.IndexOf('|');
            var second = first < 0 ? -1 : key.IndexOf('|', first + 1);
            SemanticFieldKind kind;
            if (first < 1 || second <= first || !Enum.TryParse(key.Substring(0, first), out kind)) return null;
            return new SemanticField
            {
                Kind = kind,
                Table = key.Substring(first + 1, second - first - 1),
                Name = key.Substring(second + 1)
            };
        }

        private static string SafeNameToken(string value) => (value ?? "").Replace('-', '_').Replace(' ', '_');
        private static string StatePartName(string tableName, int index) =>
            Prefix + SafeNameToken(tableName) + "_Part_" + index;

        private static int ParseChunkCount(string header)
        {
            int count;
            return header != null && header.StartsWith("b64:", StringComparison.OrdinalIgnoreCase) &&
                   int.TryParse(header.Substring(4), out count) ? count : 0;
        }

        private static string LoadStringName(Excel.Workbook workbook, string nameText)
        {
            try
            {
                var name = workbook.Names.Item(nameText);
                return Convert.ToString(name.RefersTo).TrimStart('=').Trim('"').Replace("\"\"", "\"");
            }
            catch { return null; }
        }

        private static void SaveStringName(Excel.Workbook workbook, string nameText, string value)
        {
            try { workbook.Names.Item(nameText).Delete(); } catch { }
            var escaped = (value ?? "").Replace("\"", "\"\"");
            var name = workbook.Names.Add(nameText, "=\"" + escaped + "\"");
            name.Visible = false;
        }
    }
}
