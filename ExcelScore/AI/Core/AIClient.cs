using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelScore.AI.Core
{
    public class AIClient
    {
        private readonly HttpClient _httpClient;

        public AIClient()
        {
            string apiKey =
                Environment.GetEnvironmentVariable("OMNIROUTE_API_KEY");

            if (string.IsNullOrWhiteSpace(apiKey))
                throw new Exception("OMNIROUTE_API_KEY was not found.");

            _httpClient = new HttpClient();

            _httpClient.BaseAddress =
                new Uri("http://localhost:20128/v1/");

            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", apiKey);
        }


        public async Task<AIResponse> SendAsync(Conversation conversation)
        {
            var apiRequest = new
            {
                model = "auto",
                messages = conversation.Messages.Select(message => new
                {
                    role = message.Role,
                    content = message.Content
                }).ToArray()
            };

            string json = JsonSerializer.Serialize(apiRequest);

            using (var content = new StringContent(
                json,
                Encoding.UTF8,
                "application/json"))
            {
                HttpResponseMessage response;

                try
                {
                    response =
                        await _httpClient.PostAsync(
                            "chat/completions",
                            content);
                }
                catch (HttpRequestException ex)
                {
                    throw new Exception(
                        "Could not connect to the AI service. " +
                        "Please make sure OmniRoute is running.",
                        ex);
                }

                string responseText =
                    await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception(
                        $"AI request failed: {(int)response.StatusCode}\n" +
                        responseText);
                }

                using (JsonDocument document =
                    JsonDocument.Parse(responseText))
                {
                    string answer =
                        document.RootElement
                            .GetProperty("choices")[0]
                            .GetProperty("message")
                            .GetProperty("content")
                            .GetString();

                    string model =
                        document.RootElement
                            .GetProperty("model")
                            .GetString();

                    int promptTokens =
                        document.RootElement
                            .GetProperty("usage")
                            .GetProperty("prompt_tokens")
                            .GetInt32();

                    int completionTokens =
                        document.RootElement
                            .GetProperty("usage")
                            .GetProperty("completion_tokens")
                            .GetInt32();

                    int totalTokens =
                        document.RootElement
                            .GetProperty("usage")
                            .GetProperty("total_tokens")
                            .GetInt32();

                    return new AIResponse
                    {
                        Content = answer,
                        Model = model,
                        PromptTokens = promptTokens,
                        CompletionTokens = completionTokens,
                        TotalTokens = totalTokens
                    };
                }
            }
        }

        
    }
}