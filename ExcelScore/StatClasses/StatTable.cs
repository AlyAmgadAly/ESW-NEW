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

        public string TableName { get; set; }

        public List<Parameter> Parameters { get; set; }


        public abstract string TableType { get; }


        // Methods

        //Get GroupParameters
        public List<Parameter> GetGroupParameters()
        {
            return Parameters.Where(p => p.IsGroup).ToList();
        }


    }
}
