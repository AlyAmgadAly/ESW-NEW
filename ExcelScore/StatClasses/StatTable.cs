using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.StatClasses
{
    public abstract class StatTable
    {
        // --- Basic Info ---
        public string TableName { get; set; }

        // --- Parameters ---
        public List<StatParameter> Parameters { get; set; } = new();

        // --- Abstract Type to be overridden (e.g., "Comparative", "Relation") ---
        public abstract string TableType { get; }

        public string PercentageMode { get; set; } = "Column"; 

        // --- Utility Methods ---

        // Return all grouping parameters
        public List<StatParameter> GetGroupParameters()
        {
            return Parameters.Where(p => p.IsGroup).ToList();
        }

        // Return all non-group parameters
        public List<StatParameter> GetNonGroupParameters()
        {
            return Parameters.Where(p => !p.IsGroup).ToList();
        }

        // Get a parameter by name
        public StatParameter GetParameterByName(string name)
        {
            return Parameters.FirstOrDefault(p => p.Name == name);
        }

        // Assign grouping for all non-group parameters
        public void AssignGroupedValuesToAll()
        {
            var groupParams = GetGroupParameters();
            if (!groupParams.Any()) return;

            foreach (var param in GetNonGroupParameters())
            {
                param.AssignGroupedValues(groupParams);
            }
        }
    }
}
