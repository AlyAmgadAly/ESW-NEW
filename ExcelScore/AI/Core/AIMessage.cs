using System;

namespace ExcelScore.AI.Core
{
    public class AIMessage
    {
        public string Role { get; set; }

        public string Content { get; set; }

        public DateTime Timestamp { get; set; }
    }
}