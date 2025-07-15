using System.Collections.Generic;

namespace ExcelScore.StatClasses
{
    public class RelationStatTable : StatTable
    {
        public override string TableType => "Relation";

        public RelationStatTable(string tableName, List<StatParameter> parameters)
        {
            TableName = tableName;
            Parameters = parameters;
        }

        public RelationStatTable() { }
    }
}
