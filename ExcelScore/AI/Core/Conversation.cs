using System;
using System.Collections.Generic;

namespace ExcelScore.AI.Core
{
    public class Conversation
    {
        public string Id { get; set; }

        public DateTime CreatedAt { get; set; }

        public List<AIMessage> Messages { get; set; }

        public Conversation()
        {
            Id = Guid.NewGuid().ToString();
            CreatedAt = DateTime.Now;
            Messages = new List<AIMessage>();
        }

        public void AddUserMessage(string content)
        {
            Messages.Add(new AIMessage
            {
                Role = "user",
                Content = content,
                Timestamp = DateTime.Now
            });
        }

        public void AddAssistantMessage(string content)
        {
            Messages.Add(new AIMessage
            {
                Role = "assistant",
                Content = content,
                Timestamp = DateTime.Now
            });
        }
    }
}