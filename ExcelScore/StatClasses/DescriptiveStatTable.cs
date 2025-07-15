using System.Collections.Generic;

namespace ExcelScore.StatClasses
{
    public class DescriptiveStatTable : StatTable
    {
        public override string TableType => "Descriptive";

        public DescriptiveStatTable(string tableName, List<StatParameter> parameters)
        {
            TableName = tableName;
            Parameters = parameters;
        }

        public DescriptiveStatTable() { }
    }
}
