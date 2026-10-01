using ExcelScore.AI.Core;
using ExcelScore.AI.Models;
using ExcelScore.StatClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelScore.AI.Services
{
    public class ParameterCategorizationService
    {
        
        private readonly ConversationService _conversationService;
        public ParameterCategorizationService(
    ConversationService conversationService)
        {
            _conversationService = conversationService;
        }

        private string BuildCategorizationPrompt(
    List<StatParameter> parameters)
        {
            string parameterNames = string.Join(
                ", ",
                parameters
                    .Where(p => !string.IsNullOrWhiteSpace(p.Name))
                    .Select(p => p.Name));

            string prompt =
                "You are a statistical research assistant.\n\n" +
                "Categorize the following research parameters into meaningful groups.\n\n" +
                "Rules:\n" +
                "1. Every parameter must belong to exactly one category.\n" +
                "2. Do not rename any parameter.\n" +
                "3. Do not invent new parameters.\n" +
                "4. Do not provide explanations.\n" +
                "5. Return ONLY valid JSON.\n" +
                "6. Use exactly these JSON properties: categoryName and parameters.\n\n" +
                "7. Treat parameter names as exact identifiers.\n" +
"8. Parameters with related names may belong to the same category, " +
"but each parameter remains a separate parameter.\n" +
"9. Do not combine, rename, split, or modify parameter names.\n" +
"10. Categorize based on the meaning of the parameter, " +
"not merely its name pattern.\n\n" +
                "Required JSON format:\n" +
                "[\n" +
                "  {\n" +
                "    \"categoryName\": \"Example Category\",\n" +
                "    \"parameters\": [\"Parameter 1\", \"Parameter 2\"]\n" +
                "  }\n" +
                "]\n\n" +
                "Parameters:\n" +
                parameterNames;

            return prompt;
        }

        public async Task<List<ParameterCategory>> CategorizeAsync(
    Conversation conversation,
    List<StatParameter> parameters)
        {

            if (parameters == null || parameters.Count == 0)
            {
                throw new Exception(
                    "No parameters were selected for categorization.");
            }

            string prompt =
     BuildCategorizationPrompt(parameters);

            AIResponse response =
    await _conversationService.SendMessageAsync(
        conversation,
        prompt);

            List<ParameterCategory> categories;

            try
            {
                categories =
                    JsonSerializer.Deserialize<List<ParameterCategory>>(
                        response.Content,
                        new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true
                        });
            }
            catch (JsonException ex)
            {
                throw new Exception(
                    "AI returned an invalid categorization response.",
                    ex);
            }

            if (categories == null || categories.Count == 0)
            {
                throw new Exception(
                    "AI returned an empty categorization result.");
            }

            ValidateCategorization(categories, parameters);

            return categories;
        }

        public async Task<List<ParameterCategory>>
    UpdateCategorizationAsync(
        Conversation conversation,
        List<StatParameter> parameters,
        string instruction)
        {
            AIResponse response =
                await _conversationService.SendMessageAsync(
                    conversation,
                    instruction);

            List<ParameterCategory> categories =
                JsonSerializer.Deserialize<List<ParameterCategory>>(
                    response.Content,
                    new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });

            ValidateCategorization(
    categories,
    parameters);

            return categories;
        }
        
        private void ValidateCategorization(
    List<ParameterCategory> categories,
    List<StatParameter> selectedParameters)
        {

            if (selectedParameters == null || selectedParameters.Count == 0)
            {
                throw new Exception(
                    "No parameters were provided for categorization validation.");
            }
            if (categories == null || categories.Count == 0)
            {
                throw new Exception(
                    "AI categorization returned no categories.");
            }

            foreach (ParameterCategory category in categories)
            {
                if (category == null)
                {
                    throw new Exception(
                        "AI categorization returned an invalid category.");
                }

                if (string.IsNullOrWhiteSpace(category.CategoryName))
                {
                    throw new Exception(
                        "AI categorization returned a category with no name.");
                }

                if (category.Parameters == null)
                {
                    throw new Exception(
                        "AI categorization returned a category with no parameters.");
                }
            }

            List<string> expectedParameters =
    selectedParameters
        .Where(p =>
            p != null &&
            !string.IsNullOrWhiteSpace(p.Name))
        .Select(p => p.Name.Trim())
        .ToList();

            List<string> returnedParameters =
                categories
                    .SelectMany(c => c.Parameters)
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Select(p => p.Trim())
                    .ToList();

            HashSet<string> expectedSet =
                new HashSet<string>(
                    expectedParameters,
                    StringComparer.OrdinalIgnoreCase);

            HashSet<string> returnedSet =
                new HashSet<string>(
                    returnedParameters,
                    StringComparer.OrdinalIgnoreCase);

            List<string> missingParameters =
                expectedSet
                    .Except(returnedSet)
                    .ToList();

            List<string> unexpectedParameters =
                returnedSet
                    .Except(expectedSet)
                    .ToList();

            List<string> duplicateParameters =
                returnedParameters
                    .GroupBy(
                        p => p,
                        StringComparer.OrdinalIgnoreCase)
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();

            if (missingParameters.Count > 0 ||
                unexpectedParameters.Count > 0 ||
                duplicateParameters.Count > 0)
            {
                string error =
                    "AI categorization validation failed.";

                if (missingParameters.Count > 0)
                {
                    error +=
                        "\n\nMissing parameters:\n" +
                        string.Join("\n", missingParameters);
                }

                if (unexpectedParameters.Count > 0)
                {
                    error +=
                        "\n\nUnexpected parameters:\n" +
                        string.Join("\n", unexpectedParameters);
                }

                if (duplicateParameters.Count > 0)
                {
                    error +=
                        "\n\nDuplicate parameters:\n" +
                        string.Join("\n", duplicateParameters);
                }

                throw new Exception(error);
            }
        }


    }
}