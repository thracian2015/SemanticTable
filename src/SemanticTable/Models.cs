using System.Collections.Generic;

namespace SemanticTable
{
    internal enum SemanticFieldKind { Column, Measure }

    internal sealed class SemanticField
    {
        public string Table { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public string DisplayFolder { get; set; }
        public string SortByTable { get; set; }
        public string SortByColumn { get; set; }
        internal long MetadataId { get; set; }
        internal long SortByColumnId { get; set; }
        internal bool IsHidden { get; set; }
        public SemanticFieldKind Kind { get; set; }
        public string Display => Name;
        public string Key => $"{Kind}|{Table}|{Name}";
    }

    internal sealed class SemanticHierarchy
    {
        public string Table { get; set; }
        public string Name { get; set; }
        public string DisplayFolder { get; set; }
        public List<SemanticField> Levels { get; set; } = new List<SemanticField>();
    }

    internal sealed class FieldFilter
    {
        public SemanticField Field { get; set; }
        public string Mode { get; set; } = "Basic";
        public string Operator { get; set; } = "Equals";
        public string Value { get; set; }
        public string Value2 { get; set; }
        public List<string> Values { get; set; } = new List<string>();
    }

    internal sealed class TableDefinition
    {
        public int Version { get; set; } = 5;
        public string DatasetId { get; set; }
        public string ExcelTableName { get; set; }
        public List<SemanticField> Fields { get; set; } = new List<SemanticField>();
        public List<FieldFilter> Filters { get; set; } = new List<FieldFilter>();
        public int RowLimit { get; set; } = 500000;
        public bool DeferUpdate { get; set; } = true;
    }
}
