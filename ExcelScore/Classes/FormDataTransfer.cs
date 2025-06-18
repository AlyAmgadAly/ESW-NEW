using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{

    public static class FormDataTransfer
    {
        private static Dictionary<string, object> _data = new Dictionary<string, object>();

        /// <summary>
        /// Stores a key-value pair of any data type.
        /// </summary>
        public static void Set<T>(string key, T value)
        {
            _data[key] = value;
        }

        /// <summary>
        /// Retrieves data of any type. Returns default if key not found or wrong type.
        /// </summary>
        public static T Get<T>(string key)
        {
            if (_data.ContainsKey(key) && _data[key] is T typedValue)
            {
                return typedValue;
            }

            return default;
        }

        /// <summary>
        /// Removes a specific key.
        /// </summary>
        public static void Remove(string key)
        {
            if (_data.ContainsKey(key))
            {
                _data.Remove(key);
            }
        }

        /// <summary>
        /// Clears all transferred data.
        /// </summary>
        public static void Clear()
        {
            _data.Clear();
        }
    }

}
