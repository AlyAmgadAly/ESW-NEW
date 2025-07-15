using System.Collections.Generic;

namespace ExcelScore.StatClasses
{
    public class ComparativeStatTable : StatTable
    {
        // --- Override Table Type ---
        public override string TableType => "Comparative";

        // --- Constructor (optional setup) ---
        public ComparativeStatTable(string tableName, List<StatParameter> parameters)
        {
            TableName = tableName;
            Parameters = parameters;
        }

        public ComparativeStatTable() { }
    }
}
