using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace SemanticTable
{
    // Parses the consumer schema used by Excel. Captions may be translated; DAX
    // references must use the model identifiers in the unique names instead.
    internal static class ConsumerMetadata
    {
        internal static List<string> IdentifierParts(string value) => Regex.Matches(value ?? "", @"\[((?:[^\]]|\]\])*)\]")
            .Cast<Match>().Select(m => m.Groups[1].Value.Replace("]]", "]")).ToList();

        private static string Text(DataRow row, string name) => row.Table.Columns.Contains(name) && !row.IsNull(name)
            ? Convert.ToString(row[name]) : string.Empty;
        private static int Number(DataRow row, string name) => int.TryParse(Text(row, name), out var n) ? n : 0;
        private static bool Visible(DataRow row, string name) => !row.Table.Columns.Contains(name) || row.IsNull(name)
            || Convert.ToBoolean(row[name]);
        private static string TableName(DataRow row) => IdentifierParts(Text(row, "DIMENSION_UNIQUE_NAME")).FirstOrDefault();

        internal static IReadOnlyList<SemanticField> Parse(DataTable attributes, DataTable levels, DataTable measures,
            out IReadOnlyList<SemanticHierarchy> hierarchies)
        {
            var fields = new Dictionary<string, SemanticField>(StringComparer.Ordinal);
            var byAttribute = new Dictionary<string, SemanticField>(StringComparer.Ordinal);
            foreach (DataRow row in attributes.Rows)
            {
                if (!Visible(row, "DIMENSION_IS_VISIBLE") || !Visible(row, "HIERARCHY_IS_VISIBLE")) continue;
                var origin = Number(row, "HIERARCHY_ORIGIN");
                if ((origin & 1) != 0 || (origin & (2 | 4 | 8)) == 0) continue;
                var unique = Text(row, "HIERARCHY_UNIQUE_NAME");
                var parts = IdentifierParts(unique);
                var table = TableName(row);
                if (string.IsNullOrEmpty(table) || parts.Count < 2) continue;
                var level = levels.Rows.Cast<DataRow>().FirstOrDefault(l =>
                    Text(l, "HIERARCHY_UNIQUE_NAME") == unique && (Number(l, "LEVEL_TYPE") & 1) == 0);
                var field = new SemanticField { Table = table, Name = parts[1], Kind = SemanticFieldKind.Column,
                    DisplayFolder = Text(row, "HIERARCHY_DISPLAY_FOLDER"),
                    DataType = level == null ? "" : DbTypeName(Number(level, "LEVEL_DBTYPE")) };
                fields[field.Key] = field;
                byAttribute[unique] = field;
            }
            foreach (DataRow row in measures.Rows)
            {
                if (!Visible(row, "MEASURE_IS_VISIBLE")) continue;
                // Tabular explicit measures are calculated members. Cube-generated
                // implicit aggregates do not have a corresponding DAX measure.
                if (row.Table.Columns.Contains("MEASURE_AGGREGATOR") && !row.IsNull("MEASURE_AGGREGATOR")
                    && Number(row, "MEASURE_AGGREGATOR") != 127) continue;
                var parts = IdentifierParts(Text(row, "MEASURE_UNIQUE_NAME"));
                var name = parts.Count >= 2 ? parts.Last() : Text(row, "MEASURE_NAME");
                var table = Text(row, "MEASUREGROUP_NAME");
                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(table)) continue;
                var field = new SemanticField { Table = table, Name = name, Kind = SemanticFieldKind.Measure,
                    DisplayFolder = Text(row, "MEASURE_DISPLAY_FOLDER"), DataType = DbTypeName(Number(row, "DATA_TYPE")) };
                fields[field.Key] = field;
            }
            var result = new List<SemanticHierarchy>();
            foreach (DataRow row in attributes.Rows)
            {
                if ((Number(row, "HIERARCHY_ORIGIN") & 1) == 0 || !Visible(row, "DIMENSION_IS_VISIBLE")
                    || !Visible(row, "HIERARCHY_IS_VISIBLE")) continue;
                var unique = Text(row, "HIERARCHY_UNIQUE_NAME");
                var parts = IdentifierParts(unique);
                var table = TableName(row);
                if (parts.Count < 2 || string.IsNullOrEmpty(table)) continue;
                var hierarchy = new SemanticHierarchy { Table = table, Name = parts[1],
                    DisplayFolder = Text(row, "HIERARCHY_DISPLAY_FOLDER") };
                var complete = true;
                foreach (var level in levels.Rows.Cast<DataRow>().Where(l => Text(l, "HIERARCHY_UNIQUE_NAME") == unique
                    && (Number(l, "LEVEL_TYPE") & 1) == 0).OrderBy(l => Number(l, "LEVEL_NUMBER")))
                {
                    var attribute = Text(level, "LEVEL_ATTRIBUTE_HIERARCHY_NAME");
                    // Providers may return either a full unique name or its local name.
                    if (!byAttribute.TryGetValue(attribute, out var field))
                    {
                        var names = IdentifierParts(attribute);
                        var column = names.Count > 0 ? names.Last() : attribute;
                        field = fields.Values.FirstOrDefault(f => f.Kind == SemanticFieldKind.Column
                            && f.Table == table && f.Name == column);
                    }
                    if (!Visible(level, "LEVEL_IS_VISIBLE") || field == null) { complete = false; break; }
                    hierarchy.Levels.Add(field);
                }
                // Do not substitute a level caption: it need not be a DAX column name.
                if (complete && hierarchy.Levels.Count > 0) result.Add(hierarchy);
            }
            hierarchies = result.GroupBy(h => h.Table + "\0" + h.Name).Select(g => g.First()).ToList();
            return fields.Values.ToList();
        }

        internal static string DbTypeName(int type)
        {
            switch (type)
            {
                case 11: return "Boolean";
                case 7: case 133: case 134: case 135: return "DateTime";
                case 2: case 3: case 16: case 17: case 18: case 19: case 20: case 21: return "Int64";
                case 4: case 5: return "Double";
                case 6: case 14: case 131: return "Decimal";
                case 8: case 129: case 130: return "String";
                default: return "";
            }
        }
    }
}
