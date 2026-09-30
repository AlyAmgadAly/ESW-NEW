using ExcelScore.AI.Core;
using ExcelScore.AI.Models;
using ExcelScore.AI.Services;
using ExcelScore.Classes;
using ExcelScore.StatClasses;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelScore.StatClasses.SpssReaderStat;



namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    public partial class CategorizeFrm_Child : Form
    {
        public List<StatParameter> SpssParameters;
        private readonly AIClient _aiClient;
        private readonly ParameterCategorizationService _categorizationService;
        private readonly ConversationService _conversationService;
        public CategorizeFrm_Child()
        {
            InitializeComponent();
            _aiClient = new AIClient();

            _conversationService =
                new ConversationService(_aiClient);

            _categorizationService =
    new ParameterCategorizationService(
        _conversationService);
        }

        private async void btnTestAI_Click(object sender, EventArgs e)
        {
            try
            {
                Conversation conversation =
    _conversationService.CreateConversation();


                List<ParameterCategory> categories =
                    await _categorizationService.CategorizeAsync(
                        conversation,
                        SpssParameters);

                
                string result = "";

                foreach (ParameterCategory category in categories)
                {
                    result += category.CategoryName + "\n";

                    foreach (string parameter in category.Parameters)
                    {
                        result += "  - " + parameter + "\n";
                    }

                    result += "\n";
                }

                MessageBox.Show(result, "C# Categories");

                List<ParameterCategory> updatedCategories =
    await _categorizationService.UpdateCategorizationAsync(
        conversation,
        SpssParameters,
        "Move BMI to Demographic Characteristics. " +
        "Return the complete updated categorization as JSON only.");

                string updatedResult = "";

                foreach (ParameterCategory category in updatedCategories)
                {
                    updatedResult += category.CategoryName + "\n";

                    foreach (string parameter in category.Parameters)
                    {
                        updatedResult += "  - " + parameter + "\n";
                    }

                    updatedResult += "\n";
                }

                MessageBox.Show(
                    updatedResult,
                    "Updated Categories");

            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.ToString(),
                    "AI Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void CategorizeFrm_Child_Load(object sender, EventArgs e)
        {
            SpssParameters = FormDataTransfer.Get<List<StatParameter>>("SPSS_Parameters");
        }
    }
}
