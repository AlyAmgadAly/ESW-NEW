using SpssLib.DataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.StatClasses
{
    public class SpssReaderStat
    {
        public class SpssParameter
        {
            public string Name { get; set; }
            public string Label { get; set; }
            public string Type { get; set; } // "Nominal" or "Scale"
            public List<string> Values { get; set; } = new();
            public Dictionary<int, string> ValueLabels { get; set; } = new();
        }


        public static List<SpssParameter> ReadSpssParameters(string filePath)
        {
            var result = new List<SpssParameter>();

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            using (SpssReader spssDataset = new SpssReader(fileStream))
            {
                var variables = spssDataset.Variables.ToList();

                // Prepare result for each variable
                foreach (var variable in variables)
                {
                    var param = new SpssParameter
                    {
                        Name = variable.Name,
                        Label = variable.Label,
                        Type = variable.MeasurementType.ToString().IndexOf("Nominal", StringComparison.OrdinalIgnoreCase) >= 0 ? "Nominal" : "Scale",
                        ValueLabels = new Dictionary<int, string>(),
                        Values = new List<string>()
                    };

                    // Add coded labels
                    foreach (var kvp in variable.ValueLabels)
                    {
                        if (kvp.Key is double d && d == (int)d)
                            param.ValueLabels[(int)d] = kvp.Value;
                    }

                    result.Add(param);
                }

                // Read Records
                foreach (var record in spssDataset.Records)
                {
                    for (int i = 0; i < result.Count; i++)
                    {
                        var value = record.GetValue(variables[i]);
                        string strValue = value?.ToString() ?? ".";
                        result[i].Values.Add(strValue);
                    }
                }

                // Post-process Nominal variables: ensure all values are labeled
                foreach (var param in result)
                {
                    if (param.Type == "Nominal")
                    {
                        var distinctInts = param.Values
                            .Where(v => int.TryParse(v, out _))
                            .Select(v => int.Parse(v))
                            .Distinct()
                            .OrderBy(v => v)
                            .ToList();

                        foreach (int val in distinctInts)
                        {
                            if (!param.ValueLabels.ContainsKey(val))
                            {
                                param.ValueLabels[val] = val.ToString(); // Assign raw value as label
                            }
                        }

                        // Optional: sort ValueLabels by key
                        param.ValueLabels = param.ValueLabels
                            .OrderBy(kv => kv.Key)
                            .ToDictionary(kv => kv.Key, kv => kv.Value);
                    }
                }
            }

            return result;
        }

    }
}
