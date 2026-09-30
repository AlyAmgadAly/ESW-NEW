using System.Threading.Tasks;
using ExcelScore.AI.Core;

namespace ExcelScore.AI.Services
{
    public class ConversationService
    {
        private readonly AIClient _aiClient;

        public ConversationService(AIClient aiClient)
        {
            _aiClient = aiClient;
        }

        public Conversation CreateConversation()
        {
            return new Conversation();
        }

        public async Task<AIResponse> SendMessageAsync(
            Conversation conversation,
            string message)
        {
            conversation.AddUserMessage(message);

            AIResponse response =
                await _aiClient.SendAsync(conversation);

            conversation.AddAssistantMessage(
                response.Content);

            return response;
        }
    }
}