namespace ExcelScore.AI.Core
{
    public class AIResponse
    {
        public string Content { get; set; }

        public string Model { get; set; }

        public int PromptTokens { get; set; }

        public int CompletionTokens { get; set; }

        public int TotalTokens { get; set; }
    }
}