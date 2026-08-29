using System;
using System.Collections.Generic;
using System.Linq;

namespace SemanticTable
{
    internal static class DaxQueryBuilder
    {
        public static List<SemanticField> FindReferencedFields(string dax, IEnumerable<SemanticField> availableFields)
        {
            var query = OutputProjectionText(dax ?? string.Empty);
            return availableFields.Where(field =>
            {
                if (field.Kind == SemanticFieldKind.Column)
                    return query.IndexOf(ColumnReference(field), StringComparison.OrdinalIgnoreCase) >= 0;
                return query.IndexOf(MeasureReference(field), StringComparison.OrdinalIgnoreCase) >= 0 ||
                       query.IndexOf(UnqualifiedMeasureReference(field), StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();
        }

        public static string Build(IReadOnlyCollection<SemanticField> fields, IReadOnlyCollection<FieldFilter> filters, int rowLimit)
        {
            if (fields == null || fields.Count == 0)
                throw new InvalidOperationException("Select at least one field.");

            rowLimit = Math.Max(1, Math.Min(rowLimit, 1000000));
            var columns = fields.Where(f => f.Kind == SemanticFieldKind.Column).ToList();
            var measures = fields.Where(f => f.Kind == SemanticFieldKind.Measure).ToList();
            var activeFilters = (filters ?? Array.Empty<FieldFilter>()).Where(HasCondition).ToList();
            var filterVariables = activeFilters.Select((filter, index) => new
            {
                Name = "__DS0FilterTable" + (index == 0 ? string.Empty : (index + 1).ToString()),
                Expression = FilterExpression(filter)
            }).ToList();

            var coreArguments = new List<string>();
            coreArguments.AddRange(columns.Select(ColumnReference));
            coreArguments.AddRange(filterVariables.Select(f => f.Name));
            coreArguments.AddRange(measures.Select(f => $"{DaxString(f.Name)}, {MeasureReference(f)}"));

            var define = "DEFINE\r\n";
            foreach (var filter in filterVariables)
                define += "    VAR " + filter.Name + " =\r\n" + Indent(filter.Expression, 8) + "\r\n";
            define += "    VAR __DS0Core =\r\n" +
                      "        SUMMARIZECOLUMNS (\r\n            " +
                      string.Join(",\r\n            ", coreArguments) +
                      "\r\n        )\r\n";

            var sortExpression = columns.Count > 0 ? ColumnReference(columns[0]) : MeasureReference(measures[0]);
            define += "    VAR __DS0BodyLimited =\r\n" +
                      "        TOPN ( " + rowLimit + ", __DS0Core, " + sortExpression + ", 1 )\r\n\r\n";
            return define + "EVALUATE\r\n__DS0BodyLimited\r\nORDER BY " + sortExpression;
        }

        private static string ColumnReference(SemanticField f) =>
            $"'{f.Table.Replace("'", "''")}'[{f.Name.Replace("]", "]]")}]";

        private static string FilterExpression(FieldFilter filter)
        {
            var column = ColumnReference(filter.Field);
            var value = DaxLiteral(filter.Field, filter.Value);
            if (filter.Operator == "Equals" && filter.Mode != "Advanced")
            {
                var selectedValues = filter.Values != null && filter.Values.Count > 0
                    ? filter.Values
                    : string.IsNullOrWhiteSpace(filter.Value) ? new List<string>() : new List<string> { filter.Value };
                if (selectedValues.Count == 0)
                    throw new InvalidOperationException("Select at least one value for " + filter.Field.Table + "[" + filter.Field.Name + "].");
                return "TREATAS ( { " +
                       string.Join(", ", selectedValues.Select(v => DaxLiteral(filter.Field, v))) +
                       " }, " + column + " )";
            }
            if (filter.Operator == "Equals")
                return "TREATAS ( { " + value + " }, " + column + " )";
            string condition;
            switch (filter.Operator)
            {
                case "Not Equals": condition = column + " <> " + value; break;
                case "Before": condition = column + " < " + value; break;
                case "After": condition = column + " > " + value; break;
                case "Greater Than": condition = column + " > " + value; break;
                case "Greater Than Or Equal": condition = column + " >= " + value; break;
                case "Less Than": condition = column + " < " + value; break;
                case "Less Than Or Equal": condition = column + " <= " + value; break;
                case "Between": condition = column + " >= " + value + " && " + column + " <= " + DaxLiteral(filter.Field, filter.Value2); break;
                case "Contains": condition = "CONTAINSSTRING(" + column + ", " + value + ")"; break;
                case "Starts With": condition = "LEFT(" + column + ", LEN(" + value + ")) = " + value; break;
                case "Ends With": condition = "RIGHT(" + column + ", LEN(" + value + ")) = " + value; break;
                default: condition = column + " = " + value; break;
            }
            return "FILTER ( KEEPFILTERS ( VALUES ( " + column + " ) ), " + condition + " )";
        }

        internal static bool HasCondition(FieldFilter filter)
        {
            if (filter?.Field == null) return false;
            if (string.Equals(filter.Mode, "Basic", StringComparison.OrdinalIgnoreCase))
                return (filter.Values != null && filter.Values.Count > 0) || !string.IsNullOrWhiteSpace(filter.Value);
            if (string.IsNullOrWhiteSpace(filter.Value)) return false;
            return !string.Equals(filter.Operator, "Between", StringComparison.OrdinalIgnoreCase) ||
                   !string.IsNullOrWhiteSpace(filter.Value2);
        }

        internal static string ConditionSignature(FieldFilter filter) =>
            HasCondition(filter) ? FilterExpression(filter) : string.Empty;

        private static string Indent(string value, int spaces)
        {
            var prefix = new string(' ', spaces);
            return prefix + value.Replace("\r\n", "\r\n" + prefix);
        }

        private static string OutputProjectionText(string dax)
        {
            const string function = "SUMMARIZECOLUMNS";
            var start = dax.IndexOf(function, StringComparison.OrdinalIgnoreCase);
            if (start < 0) return dax;
            start = dax.IndexOf('(', start + function.Length);
            if (start < 0) return dax;
            var depth = 0;
            for (var i = start; i < dax.Length; i++)
            {
                if (dax[i] == '(') depth++;
                else if (dax[i] == ')' && --depth == 0) return dax.Substring(start, i - start + 1);
            }
            return dax;
        }

        private static string DaxLiteral(SemanticField field, string value)
        {
            value = value ?? string.Empty;
            DateTime date;
            if (IsDate(field) && DateTime.TryParse(value, out date))
                return $"DATE({date.Year}, {date.Month}, {date.Day})";
            decimal number;
            if (IsNumeric(field) && decimal.TryParse(value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out number))
                return number.ToString(System.Globalization.CultureInfo.InvariantCulture);
            bool boolean;
            if (IsBoolean(field) && bool.TryParse(value, out boolean)) return boolean ? "TRUE()" : "FALSE()";
            return DaxString(value);
        }

        internal static bool IsDate(SemanticField field)
        {
            var type = field?.DataType ?? string.Empty;
            var name = field?.Name ?? string.Empty;
            return type == "9" || type.IndexOf("date", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("date", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("month start", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        internal static bool IsNumeric(SemanticField field)
        {
            var type = field?.DataType ?? string.Empty;
            return type == "6" || type == "8" || type == "10" ||
                   type.IndexOf("int", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   type.IndexOf("decimal", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   type.IndexOf("double", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        internal static bool IsBoolean(SemanticField field)
        {
            var type = field?.DataType ?? string.Empty;
            return type == "1" || type.IndexOf("bool", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string MeasureReference(SemanticField f) =>
            $"'{f.Table.Replace("'", "''")}'[{f.Name.Replace("]", "]]")}]";
        private static string UnqualifiedMeasureReference(SemanticField f) =>
            $"[{f.Name.Replace("]", "]]")}]";
        private static string DaxString(string value) => $"\"{value.Replace("\"", "\"\"")}\"";
    }
}
