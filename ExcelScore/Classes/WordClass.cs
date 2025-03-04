using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Humanizer;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Aspose.Cells;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.CompilerServices;
using CenterSpace.NMath.Core;
using System.Collections.ObjectModel;
using Color = Syncfusion.Drawing.Color;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using DocumentFormat.OpenXml.Spreadsheet;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using System.Collections;
using Accord.IO;
using static Humanizer.On;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ExcelScore.Classes
{
    public class WordClass
    {
        private WordDocument document;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;



        public WordDocument InitWord()
        {
            document = new WordDocument();
            
            

            return document;
        }

        public void documentformatpage()
        {

        }
        public void removeHeader(string filepath)
        {
            
            try
            {
                
                string textToFind1 = "Created with a trial version of Syncfusion Word library or registered the wrong key in your application.";
                string textToReplace = "";
                string textToFind2 = "to obtain the valid key.";
                string textToFind3 = "Click";

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepath, true))
                {
                    var mainPart = wordDoc.MainDocumentPart;


                    //removing text 1 & 2
                    // Loop through all paragraphs in the document body
                    foreach (var paragraph in mainPart.Document.Body.Elements<Paragraph>())
                    {
                        // Loop through each run in the paragraph
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            // Loop through each text element in the run
                            var textElements = run.Elements<Text>().ToList();
                            foreach (var text in textElements)
                            {
                                // Replace the text if it matches the target string
                                if (text.Text.Contains(textToFind1))
                                {
                                    text.Text = text.Text.Replace(textToFind1, textToReplace);
                                }
                                if (text.Text.Contains(textToFind2))
                                {
                                    text.Text = text.Text.Replace(textToFind2, textToReplace);
                                }
                                if (text.Text.Contains(textToFind3))
                                {
                                    text.Text = text.Text.Replace(textToFind3, textToReplace);
                                }
                            }
                        }
                    }


                    //removing hyperlinks
                    

                    // Remove all hyperlink relationships from the document
                    


                    //removing watermark
                    // Loop through each section in the document
                    foreach (var section in wordDoc.MainDocumentPart.Document.Body.Elements<SectionProperties>())
                    {
                        // Remove the header references from the section properties
                        section.RemoveAllChildren<HeaderReference>();
                    }

                    // Collect all header parts in a list first to avoid modifying the collection during iteration
                    var headerParts = wordDoc.MainDocumentPart.HeaderParts.ToList();

                    // Remove the header parts from the document
                    foreach (var header in headerParts)
                    {
                        wordDoc.MainDocumentPart.DeletePart(header);
                    }
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Iterate through all paragraphs in the body

                    // Save the document
                    wordDoc.MainDocumentPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void addparag(IWSection section , string Text)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;
            IWTextRange firstTextRange = firstParagraph.AppendText(Text);
            
        }
        public IWSection CreatePortraitSection()
        {
            IWSection Portrait = document.AddSection();
            Portrait.PageSetup.Orientation = PageOrientation.Portrait;
            //Portrait.PageSetup.Margins.Top = 
            //Margins
            //Top 3.5
            //Bottom 3
            //Left 3.25
            //Right 2.75

            //Layout
            //header 2
            //footer 2

            Portrait.PageSetup.Margins.Top = SetColumnWidthInCentimeters(3.5f);
            Portrait.PageSetup.Margins.Bottom = SetColumnWidthInCentimeters(3f); 
            Portrait.PageSetup.Margins.Left = SetColumnWidthInCentimeters(3.25f); 
            Portrait.PageSetup.Margins.Right = SetColumnWidthInCentimeters(2.75f); 

            Portrait.PageSetup.HeaderDistance = SetColumnWidthInCentimeters(2f); 
            Portrait.PageSetup.FooterDistance = SetColumnWidthInCentimeters(2f); 

            return Portrait;
        }

        public IWSection CreateLandscapeSection()
        {
            IWSection Landscape = document.AddSection();
            Landscape.PageSetup.Orientation = PageOrientation.Landscape;

            

            return Landscape;
        }

        public void AddDescriptiveTitle(IWSection section, string TableName)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;

            


            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText("Table ():	Distribution of the studied cases according to "+TableName+" in # group (n = #)");
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            //-70.56f
            paragraphFormat.FirstLineIndent = SetColumnWidthInCentimeters(-2.5f);
            paragraphFormat.LeftIndent = SetColumnWidthInCentimeters(2.5f);


            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    break;
                }
            }
        }

        //Table() :\tComparison

        public void AddCorrTitle(IWSection section, string TableName, int numberofgroups)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;

            string groupText = numberofgroups.ToWords();


            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText("Table ():\tCorrelation between different expressions in " + TableName);
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            //-70.56f
            paragraphFormat.FirstLineIndent = SetColumnWidthInCentimeters(-2.5f);
            paragraphFormat.LeftIndent = SetColumnWidthInCentimeters(2.5f);
            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    break;
                }
            }
        }
        public void AddComparativeTitle(IWSection section, string TableName , int numberofgroups)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;

            string groupText = numberofgroups.ToWords();


            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText("Table ():\tComparison between the "+ groupText + " studied groups according to "+ TableName);
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            //-70.56f
            paragraphFormat.FirstLineIndent = SetColumnWidthInCentimeters(-2.5f);
            paragraphFormat.LeftIndent = SetColumnWidthInCentimeters(2.5f);
            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    break;
                }
            }
        }

        public void AddToolTitle_Dynamic(IWSection section, string Title)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;




            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText(Title);
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;

            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    text.CharacterFormat.FontSize = 16;
                    text.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                    text.CharacterFormat.TextBackgroundColor = Syncfusion.Drawing.Color.LightBlue;
                    break;
                }
            }
        }
        public void AddTitle_Dynamic(IWSection section, string Title , ref int tableorder)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;

            


            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText("Table ("+ tableorder + "):\t"+ Title);
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            //-70.56f
            paragraphFormat.FirstLineIndent = SetColumnWidthInCentimeters(-2.5f);
            paragraphFormat.LeftIndent = SetColumnWidthInCentimeters(2.5f);
            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    break;
                }
            }
        }
        public void AddRgressionTitle(IWSection section, string TableName ,string RegressionType)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;



            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText("Table ():	Univariate and multivariate "+RegressionType+" regression analysis for the parameters affecting "+TableName);
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            //-70.56f
            paragraphFormat.FirstLineIndent = SetColumnWidthInCentimeters(-2.5f);
            paragraphFormat.LeftIndent = SetColumnWidthInCentimeters(2.5f);
            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    break;
                }
            }
        }


        public void AddRelationTitle(IWSection section, string TableName)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;

           


            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText("Table ():\tRelation between "+ TableName + " with different parameters in total sample");
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            //-70.56f
            paragraphFormat.FirstLineIndent = SetColumnWidthInCentimeters(-2.5f);
            paragraphFormat.LeftIndent = SetColumnWidthInCentimeters(2.5f);
            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    break;
                }
            }
        }
        public void AddTitle(IWSection section , string ACurrentDomainName , int datacount)
        {
            IWParagraph firstParagraph = section.AddParagraph();
            WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;

            string CurrentDomainName = ACurrentDomainName;


            //Title
            IWTextRange firstTextRange = firstParagraph.AppendText("Table ():\tDistribution of the studied patients according to " + CurrentDomainName + "  (n = " + datacount + ")");
            paragraphFormat.AfterSpacing = 10;
            paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            paragraphFormat.FirstLineIndent = -70.56f;
            paragraphFormat.LeftIndent = 70.56f;
            //Line Spacing Multiple - 1.25
            paragraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraphFormat.LineSpacing = 15f;



            foreach (ParagraphItem item in firstParagraph.ChildEntities)
            {
                if (item is WTextRange)
                {
                    WTextRange text = item as WTextRange;
                    //Modifies the character format of the text
                    text.CharacterFormat.Bold = true;
                    break;
                }
            }
        }

        public IWTable Createtable(IWSection section , int WordTableRows , int WordTableColumns)
        {
           
            IWTable table = section.AddTable();

            
            table.ResetCells(WordTableRows, WordTableColumns);
            

            return table;
        }
        public void GeneralLetterTableFormat(IWTable table)
        {
            table.TableFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThinThickSmallGap;
            table.TableFormat.Borders.Top.LineWidth = 2.25f;


            table.TableFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThinThickSmallGap;
            table.TableFormat.Borders.Left.LineWidth = 2.25f;

            table.TableFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThickThinMediumGap;
            table.TableFormat.Borders.Right.LineWidth = 2.25f;



            table.TableFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThickThinMediumGap;
            table.TableFormat.Borders.Bottom.LineWidth = 2.25f;



            

            table.TableFormat.HorizontalAlignment = RowAlignment.Center;





        }
        public void AddHeaderPaper(IWTable table, int WordTableRows , int WordTableColumns , int numberofgroups , ComparativeTable comparativeTable)
        {
            //AddPara_Center(table, 0, WordTableColumns - 2, "Test of Sig.");
            AddPara_Center(table, 0, WordTableColumns - 1, "p");


            
            int row = 1;
            foreach (var paramter in comparativeTable.Parameters)
            {
                int rowcount = 0;
                if (paramter.IsGroup)
                {
                    int col = 1;
                    foreach (var kvp in paramter.DIC_LablesIfNomainal)
                    {
                        int key = kvp.Key;
                        string label = kvp.Value;
                        int count = paramter.ParameterValues.Count(x => x == key);
                        AddPara_Center(table, 0, col, label + "\n(n = " + count + ")");
                        col++;
                    }
                }

                else if(paramter.NominalOrScale == "Nominal")
                {
                    AddPara_NoCenter(table, row, 0, paramter.Name);

                    int innerrow = row + 1;

                    int mergecountrows = paramter.DIC_LablesIfNomainal.Keys.Count;

                    table.ApplyVerticalMerge(WordTableColumns - 2, innerrow, innerrow + mergecountrows - 1);
                    table.ApplyVerticalMerge(WordTableColumns - 1, innerrow, innerrow + mergecountrows - 1);

                    foreach (var label in paramter.DIC_LablesIfNomainal.Values)
                    {
                        Addpara_NoCenterNoBOLD(table, innerrow, 0, label);
                        LeftIntendBeforeText(table, innerrow, 0, 14.17f);
                        innerrow++;

                    }



                    rowcount = paramter.DIC_LablesIfNomainal.Keys.Count + 1;
                }
                else if(paramter.NominalOrScale == "Scale")
                {
                    AddPara_NoCenter(table, row, 0, paramter.Name);


                    Addpara_NoCenterNoBOLD(table, row+1, 0, "Min. – Max.");
                    LeftIntendBeforeText(table, row + 1, 0, 14.17f);

                    Addpara_NoCenterNoBOLD(table, row+2, 0, "Mean ± SD.");
                    LeftIntendBeforeText(table, row + 2, 0, 14.17f);

                    Addpara_NoCenterNoBOLD(table, row+3, 0, "Median (IQR)");
                    LeftIntendBeforeText(table, row + 3, 0, 14.17f);


                    table.ApplyVerticalMerge(WordTableColumns - 2, row+1, row+3);
                    table.ApplyVerticalMerge(WordTableColumns - 1, row + 1, row + 3);

                    rowcount = 4;

                }



                row = row + rowcount;
            }
        }
        public void PaperComparative_widths(IWTable table , int WordTableRows , int numberofgroups,int WordTableColumns)
        {
            for(int i = 0;i<WordTableRows;i++)
            {
                table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(4f);
                table.Rows[i].Cells[WordTableColumns-2].Width = SetColumnWidthInCentimeters(1.75f);
                table.Rows[i].Cells[WordTableColumns-1].Width = SetColumnWidthInCentimeters(1.75f);


                for(int groups = 1;groups<=numberofgroups;groups++)
                {
                    table.Rows[i].Cells[groups].Width = SetColumnWidthInCentimeters(3.75f);
                }
            }
            
        }

        public void PaperBordersGeneral(IWTable table , int WordTableColumns)
        {
            for(int col = 0;col<WordTableColumns;col++)
            {
                table.Rows[0].Cells[col].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[col].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }
        }
        public void PaperGeneralFormat(IWTable table)
        {

            table.TableFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.TableFormat.Borders.Top.LineWidth = 1.5f;


            table.TableFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.TableFormat.Borders.Bottom.LineWidth = 1.5f;

            table.TableFormat.Borders.Horizontal.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Cleared;
            table.TableFormat.Borders.Vertical.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Cleared;
            table.TableFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Cleared;
            table.TableFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Cleared;
            table.TableFormat.HorizontalAlignment = RowAlignment.Center;




        }
        public void GeneralTableFormat(IWTable table)
        {
            table.TableFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThinThickSmallGap;
            table.TableFormat.Borders.Top.LineWidth = 2.25f;


            table.TableFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThinThickSmallGap;
            table.TableFormat.Borders.Left.LineWidth = 2.25f;

            table.TableFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThickThinMediumGap;
            table.TableFormat.Borders.Right.LineWidth = 2.25f;



            table.TableFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThickThinMediumGap;
            table.TableFormat.Borders.Bottom.LineWidth = 2.25f;






            table.TableFormat.Borders.Horizontal.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Cleared;

            table.TableFormat.HorizontalAlignment = RowAlignment.Center;


           


        }
        public void ApplyRelation_Widths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for(int row = 0;row < WordTableRows;row++)
            {
                table.Rows[row].Cells[0].Width = SetColumnWidthInCentimeters(3.75f);
                table.Rows[row].Cells[1].Width = SetColumnWidthInCentimeters(1f);

                table.Rows[row].Cells[2].Width = SetColumnWidthInCentimeters(3.5f);
                table.Rows[row].Cells[3].Width = SetColumnWidthInCentimeters(4.25f);

                table.Rows[row].Cells[4].Width = SetColumnWidthInCentimeters(1.75f);
                table.Rows[row].Cells[5].Width = SetColumnWidthInCentimeters(1.75f);

            }
            
        }

        public void ApplyRelation_IQR_Widths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int row = 0; row < WordTableRows; row++)
            {
                table.Rows[row].Cells[0].Width = SetColumnWidthInCentimeters(3.75f);
                table.Rows[row].Cells[1].Width = SetColumnWidthInCentimeters(1f);

                table.Rows[row].Cells[2].Width = SetColumnWidthInCentimeters(3f);
                table.Rows[row].Cells[3].Width = SetColumnWidthInCentimeters(3f);

                table.Rows[row].Cells[4].Width = SetColumnWidthInCentimeters(4f);
                table.Rows[row].Cells[WordTableColumns-2].Width = SetColumnWidthInCentimeters(1.75f);
                table.Rows[row].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.75f);

            }

        }

        public void ApplyRelation_Widths_Pathology(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int row = 0; row < WordTableRows; row++)
            {
                table.Rows[row].Cells[0].Width = SetColumnWidthInCentimeters(3.85f);
                table.Rows[row].Cells[1].Width = SetColumnWidthInCentimeters(3.75f);

                table.Rows[row].Cells[2].Width = SetColumnWidthInCentimeters(1f);

                table.Rows[row].Cells[3].Width = SetColumnWidthInCentimeters(2.75f);
                table.Rows[row].Cells[4].Width = SetColumnWidthInCentimeters(3.75f);

                table.Rows[row].Cells[5].Width = SetColumnWidthInCentimeters(1.7f);
                table.Rows[row].Cells[6].Width = SetColumnWidthInCentimeters(1.7f);

            }

        }

        public void ApplyRegression_Widths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int row = 0; row < WordTableRows; row++)
            {
                table.Rows[row].Cells[0].Width = SetColumnWidthInCentimeters(5f);
                
                table.Rows[row].Cells[1].Width = SetColumnWidthInCentimeters(1.7f);
                table.Rows[row].Cells[2].Width = SetColumnWidthInCentimeters(4.35f);

                table.Rows[row].Cells[3].Width = SetColumnWidthInCentimeters(1.7f);
                table.Rows[row].Cells[4].Width = SetColumnWidthInCentimeters(4.35f);



            }

        }

        public void InsertRelation_InnerHeader_Merges(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns)
        {
            int StartingRow = 2;

            foreach (var parameter in comparativeTable.Parameters)
            {

                
                if (parameter.IsGroup)
                {
                    int ParameterLabelCtr = parameter.DIC_LablesIfNomainal.Keys.Count;





                    AddPara_NoCenter(table, StartingRow, 0, parameter.Name);
                    StartingRow++;

                    table.ApplyVerticalMerge(WordTableColumns - 2, StartingRow, StartingRow + ParameterLabelCtr - 1);
                    table.ApplyVerticalMerge(WordTableColumns - 1, StartingRow, StartingRow + ParameterLabelCtr - 1);


                    if (StartingRow + ParameterLabelCtr - 1 < WordTableRows-1)
                    {
                        for (int col = 0; col < WordTableColumns; col++)
                        {
                            table.Rows[StartingRow + ParameterLabelCtr - 1].Cells[col].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[StartingRow + ParameterLabelCtr - 1].Cells[col].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }
                    }
                    foreach (var kvp in parameter.DIC_LablesIfNomainal)
                    {
                        int key = kvp.Key;

                        string label = kvp.Value;
                        int count = parameter.ParameterValues.Count(x => x == key);

                        if (count != 0)
                        {
                            Addpara_NoCenterNoBOLD(table, StartingRow, 0, label);
                            LeftIntendBeforeText(table, StartingRow, 0, SetColumnWidthInCentimeters(0.5f));
                            AddPara_Center(table, StartingRow, 1, count.ToString());
                            StartingRow++;
                        }
                        else if(count == 0)
                        {
                            Addpara_NoCenterNoBOLD(table, StartingRow, 0, label);
                            LeftIntendBeforeText(table, StartingRow, 0, SetColumnWidthInCentimeters(0.5f));
                            AddPara_Center(table, StartingRow, 1, "0");
                            StartingRow++;
                        }

                        
                    }
                    
                    



                }
            }
        }


        public void InsertRelation_InnerHeader_Merges_Pathology(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns)
        {
            WordClass wordObj = new WordClass();
            int StartingRow = 2;

            foreach (var parameter in comparativeTable.Parameters)
            {


                if (parameter.IsGroup)
                {
                    int ParameterLabelCtr = parameter.DIC_LablesIfNomainal.Keys.Count;

                    string InsertedN = null;
                    if (parameter.hasLowerN)
                    {
                        InsertedN = "(n = " + parameter.ParameterValues.Count + ")";

                    }

                    if (InsertedN == null)
                    {
                        AddPara_NoCenter(table, StartingRow, 0, parameter.Name);
                    }
                    if (InsertedN != null)
                    {
                        AddPara_NoCenter(table, StartingRow, 0, parameter.Name);
                        wordObj.AddParaCombined(table, StartingRow, 0, InsertedN, true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);
                    }



                    table.ApplyVerticalMerge(0, StartingRow, StartingRow + 1);

                    table.ApplyVerticalMerge(WordTableColumns - 2, StartingRow, StartingRow + ParameterLabelCtr - 1);
                    table.ApplyVerticalMerge(WordTableColumns - 1, StartingRow, StartingRow + ParameterLabelCtr - 1);


                    if (StartingRow + ParameterLabelCtr - 1 < WordTableRows - 1)
                    {
                        for (int col = 0; col < WordTableColumns; col++)
                        {
                            table.Rows[StartingRow + ParameterLabelCtr - 1].Cells[col].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[StartingRow + ParameterLabelCtr - 1].Cells[col].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }
                    }
                    foreach (var kvp in parameter.DIC_LablesIfNomainal)
                    {
                        int key = kvp.Key;

                        string label = kvp.Value;
                        int count = parameter.ParameterValues.Count(x => x == key);

                        if (count > 1)
                        {
                            Addpara_NoCenterNoBOLD(table, StartingRow, 1, label);
                            LeftIntendBeforeText(table, StartingRow, 1, SetColumnWidthInCentimeters(0.5f));
                            AddPara_Center(table, StartingRow, 2, count.ToString());
                            StartingRow++;
                        }
                        else if (count == 0)
                        {
                            Addpara_NoCenterNoBOLD(table, StartingRow, 1, label);
                            LeftIntendBeforeText(table, StartingRow, 1, SetColumnWidthInCentimeters(0.5f));
                            AddPara_Center(table, StartingRow, 2, "0");
                            StartingRow++;
                        }
                        else if (count == 1)
                        {
                            Addpara_NoCenterNoBOLD(table, StartingRow, 1, label);
                            LeftIntendBeforeText(table, StartingRow, 1, SetColumnWidthInCentimeters(0.5f));
                            AddPara_Center(table, StartingRow, 2, "1");
                            SubSuperScriptText(table, StartingRow, 2, Color.White, "#", "Super");
                            StartingRow++;
                        }


                    }





                }
            }
        }

       
        public void InsertRegression_InnerHeader_Merges(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns)
        {
            
            int StartingRow = 2;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.NominalOrScale == "Dependent")
                {
                    continue;
                }

                else if(parameter.NormalOrAbnormal == "Not Seperated")
                {

                    if(parameter.NominalOrScale == "Scale")
                    {

                        AddPara_NoCenter(table, StartingRow, 0, parameter.Name);

                        if(StartingRow < WordTableRows - 1)
                        {
                            for (int j = 0; j < WordTableColumns; j++)
                            {
                                table.Rows[StartingRow].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                                table.Rows[StartingRow].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                            }
                        }
                        

                        StartingRow++;

                    }
                    else if(parameter.NominalOrScale == "Nominal")
                    {
                        if (parameter.DIC_LablesIfNomainal[1] == "Yes")
                        {
                            AddPara_NoCenter(table, StartingRow, 0, parameter.Name+" [Yes]");
                            if (StartingRow < WordTableRows - 1)
                            {
                                for (int j = 0; j < WordTableColumns; j++)
                                {
                                    table.Rows[StartingRow].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                                    table.Rows[StartingRow].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                                }
                            }
                            StartingRow++;
                        }

                        else
                        {
                            int lastKey = parameter.DIC_LablesIfNomainal.Keys.LastOrDefault();

                            AddPara_NoCenter(table, StartingRow, 0, parameter.DIC_LablesIfNomainal[lastKey]);
                            if (StartingRow < WordTableRows - 1)
                            {
                                for (int j = 0; j < WordTableColumns; j++)
                                {
                                    table.Rows[StartingRow].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                                    table.Rows[StartingRow].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                                }
                            }
                            StartingRow++;
                        }
                    }

                }
                else if(parameter.NormalOrAbnormal == "Seperated")
                {
                    

                    int count = StartingRow + parameter.DIC_LablesIfNomainal.Keys.Count;
                    if (count < WordTableRows - 1)
                    {
                        for (int j = 0; j < WordTableColumns; j++)
                        {
                            table.Rows[count].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[count].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }
                    }

                    AddPara_NoCenter(table, StartingRow, 0, parameter.Name);
                    StartingRow++;

                    foreach (var keyValuePair in parameter.DIC_LablesIfNomainal)
                    {
                        string Value = keyValuePair.Value;
                        Addpara_NoCenterNoBOLD(table, StartingRow, 0, Value);
                        LeftIntendBeforeText(table, StartingRow, 0, SetColumnWidthInCentimeters(0.5f));
                        StartingRow++;
                    }

                }

            }
        }
        public void ApplyRelation_OuterHeaders(IWTable table,ComparativeTable comparativeTable ,int WordTableRows, int WordTableColumns)
        {
            AddPara_Center(table, 0, 1, "No.");

            AddPara_Center(table, 0, 2, comparativeTable.TableName);


            AddPara_Center(table, 1, 2, "Mean ± SD.");
            AddPara_Center(table, 1, 3, "Median (Min. – Max.)");

            AddPara_Center(table, 0, WordTableColumns - 1, "p");


        }

        public void ApplyRelation_IQR_OuterHeaders(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns)
        {
            AddPara_Center(table, 0, 1, "N");

            AddPara_Center(table, 0, 2, comparativeTable.TableName);


            AddPara_Center(table, 1, 2, "Min. – Max."); 
            AddPara_Center(table, 1, 3, "Mean ± SD.");
            AddPara_Center(table, 1, 4, "Median (IQR)");

            AddPara_Center(table, 0, WordTableColumns - 1, "p");


        }

        public void ApplyRelation_OuterHeaders_Pathology(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns)
        {
            AddPara_Center(table, 0, 2, "N");

            AddPara_Center(table, 0, 3, comparativeTable.TableName);


            AddPara_Center(table, 1, 3, "Mean ± SD.");
            AddPara_Center(table, 1, 4, "Median (Min. – Max.)");

            AddPara_Center(table, 0, WordTableColumns - 1, "p");


        }
        public void Apply_Linear_Regression_OuterHeaders(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns , string RegressionOrValue)
        {
            AddPara_Center(table, 0, 1, "Univariate");
            AddPara_Center(table, 1, 1, "p");
            AddPara_Center(table, 1, 2, RegressionOrValue + " (LL – UL 95%C.I)");

            //SubSuperScriptText(table, 0, 3, Color.White, "#", "Super");
            AddPara_Center(table, 0, 3, "Multivariate");
            AddPara_Center(table, 1, 3, "p");
            AddPara_Center(table, 1, 4, RegressionOrValue + " (LL – UL 95%C.I)");


        }


        public void ApplyRelation_OuterBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for (int k = 0; k < WordTableRows; k++)
            {
                table.Rows[k].Cells[1].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[1].CellFormat.Borders.Right.LineWidth = 1.5f;

                table.Rows[k].Cells[WordTableColumns-2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[WordTableColumns-2].CellFormat.Borders.Left.LineWidth = 1.5f;
            }

            table.Rows[0].Cells[2].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[2].CellFormat.Borders.Bottom.LineWidth = 0.5f;

            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.LineWidth = 0.5f;

        }

        public void ApplyRelation_IQR_OuterBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for (int k = 0; k < WordTableRows; k++)
            {
                table.Rows[k].Cells[1].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[1].CellFormat.Borders.Right.LineWidth = 1.5f;

                table.Rows[k].Cells[WordTableColumns - 2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[WordTableColumns - 2].CellFormat.Borders.Left.LineWidth = 1.5f;
            }

            table.Rows[0].Cells[2].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[2].CellFormat.Borders.Bottom.LineWidth = 0.5f;

            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.LineWidth = 0.5f;

            table.Rows[0].Cells[4].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[4].CellFormat.Borders.Bottom.LineWidth = 0.5f;

        }

        public void Apply_Descriptive_OuterBorders_PeriodsNoTest()
        {

        }
        public void ApplyRelation_OuterBorders_Pathology(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for (int k = 0; k < WordTableRows; k++)
            {
                table.Rows[k].Cells[2].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[2].CellFormat.Borders.Right.LineWidth = 1.5f;

                table.Rows[k].Cells[WordTableColumns - 2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[WordTableColumns - 2].CellFormat.Borders.Left.LineWidth = 1.5f;
            }

            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.LineWidth = 0.5f;

            table.Rows[0].Cells[4].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[4].CellFormat.Borders.Bottom.LineWidth = 0.5f;

        }

        public void ApplyRegression_OuterBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for (int k = 0; k < WordTableRows; k++)
            {
                table.Rows[k].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;

                table.Rows[k].Cells[3].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[3].CellFormat.Borders.Left.LineWidth = 1.5f;
            }

            table.Rows[0].Cells[1].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[1].CellFormat.Borders.Bottom.LineWidth = 0.5f;

            table.Rows[0].Cells[2].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[2].CellFormat.Borders.Bottom.LineWidth = 0.5f;

            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[3].CellFormat.Borders.Bottom.LineWidth = 0.5f;

            table.Rows[0].Cells[4].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[4].CellFormat.Borders.Bottom.LineWidth = 0.5f;

        }
        public void ApplyRelation_OuterMerges(IWTable table , int WordTableRows, int WordTableColumns)
        {
            table.ApplyVerticalMerge(1, 0, 1);
            table.ApplyVerticalMerge(WordTableColumns-2, 0, 1);
            table.ApplyVerticalMerge(WordTableColumns-1, 0, 1);

            table.ApplyHorizontalMerge(0, 2, 3);


        }

        public void ApplyRelation_IQR_OuterMerges(IWTable table, int WordTableRows, int WordTableColumns)
        {
            table.ApplyVerticalMerge(1, 0, 1);

            table.ApplyVerticalMerge(WordTableColumns - 2, 0, 1);
            table.ApplyVerticalMerge(WordTableColumns - 1, 0, 1);

            table.ApplyHorizontalMerge(0, 2, 4);


        }

        public void ApplyRelation_OuterMerges_Pathology(IWTable table, int WordTableRows, int WordTableColumns)
        {
            table.ApplyVerticalMerge(0, 0, 1);
            table.ApplyVerticalMerge(1, 0, 1);
            table.ApplyVerticalMerge(2, 0, 1);
            table.ApplyVerticalMerge(WordTableColumns - 2, 0, 1);
            table.ApplyVerticalMerge(WordTableColumns - 1, 0, 1);

            table.ApplyHorizontalMerge(0, 3, 4);


        }

        public void Apply_Descriptive_periodsNoTest_OuterMerges(IWTable table, int WordTableRows, int WordTableColumns , int numberofPeriods)
        {
            


            int startcol = 1;
            for(int i = 0; i < numberofPeriods;i++)
            {
                table.ApplyHorizontalMerge(0, startcol, startcol + 1);
                startcol = startcol +2;
            }


        }

        public void ApplyRegression_OuterMerges(IWTable table, int WordTableRows, int WordTableColumns)
        {
            table.ApplyHorizontalMerge(0, 1, 2);
            table.ApplyHorizontalMerge(0, 3, 4);


        }


        public void ApplyGeneralDescritiveBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }



            for (int k = 0; k < WordTableRows; k++)
            {
                table.Rows[k].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;
            }

            

        }
        public void ApplyGeneralLettersBorders(IWTable table, int WordTableRows, int WordTableColumns, int numberofgroups)
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;

                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;

                table.Rows[WordTableRows-2].Cells[j].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[WordTableRows - 2].Cells[j].CellFormat.Borders.Top.LineWidth = 1.5f;
            }

            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[i].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;

            }


            table.ApplyVerticalMerge(0, 0, 1);


        }
        public void ApplyGeneralMatrixCorrBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[1].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[i].Cells[1].CellFormat.Borders.Right.LineWidth = 1.5f;

            }


            //table.ApplyVerticalMerge(0, 0, 1);


        }
        public void ApplyGeneralNormalCorrBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 0; j < WordTableRows; j++)
            {
                table.Rows[j].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[j].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;
            }

            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            

        }

        public void ApplyGeneralNormalCorrMerges(IWTable table, int WordTableRows, int WordTableColumns)
        {
            table.ApplyVerticalMerge(0, 0, 1);
            for(int i = 1; i < WordTableColumns; i=i+2)
            {
                table.ApplyHorizontalMerge(0, i, i + 1);
            }

        }

        public void ApplyGeneral_Groups_2Periods_Borders(IWTable table, int WordTableRows, int WordTableColumns , ComparativeTable comparativeTable)
        {
            for(int col = 0;col< WordTableColumns;col++)
            {
                table.Rows[0].Cells[col].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[col].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for(int row = 0;row< WordTableRows;row++)
            {
                table.Rows[row].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[row].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;

                table.Rows[row].Cells[WordTableColumns-2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[row].Cells[WordTableColumns-2].CellFormat.Borders.Left.LineWidth = 1.5f;

            }



            int numberofparameters = comparativeTable.Parameters.Count - 1;

            int testrows = numberofparameters / 2;

            int rowctr = 1;

            for (int testrowctr = 1; testrowctr <= testrows; testrowctr++)
            {
                for (int col = 1; col < WordTableColumns; col++)
                {
                    table.Rows[rowctr + 8].Cells[col].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[rowctr + 8].Cells[col].CellFormat.Borders.Top.LineWidth = 0.5f;


                    
                }


                
                rowctr = rowctr + 9;

                for (int col = 0; col < WordTableColumns; col++)
                {
                    if (rowctr -1 < WordTableRows - 1)
                    {
                        table.Rows[rowctr -1].Cells[col].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[rowctr -1].Cells[col].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                    }
                }
                    
            }



                

            


        }
        public void ApplyGeneralComparativeBorders(IWTable table,  int WordTableRows, int WordTableColumns, int numberofgroups )
        {
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }



            for (int k = 0; k < WordTableRows; k++)
            {
                table.Rows[k].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[k].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;
            }

            for (int l = 0; l < WordTableRows; l++)
            {
                table.Rows[l].Cells[WordTableColumns-2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[l].Cells[WordTableColumns - 2].CellFormat.Borders.Left.LineWidth = 1.5f;

            }
            for (int i = 1; i <= numberofgroups * 2; i++)
            {
                table.Rows[1].Cells[i].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[i].CellFormat.Borders.Top.LineWidth = 0.5f;
            }

        }

        public void ApplyGeneral_PeriodsUp_Groups_ComparativeBorders(IWTable table, int WordTableRows, int WordTableColumns, int numberofgroups , ComparativeTable comparativeTable)
        {
            for(int row = 0; row < WordTableRows;row++)
            {
                table.Rows[row].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[row].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;


                table.Rows[row].Cells[WordTableColumns-2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[row].Cells[WordTableColumns-2].CellFormat.Borders.Left.LineWidth = 1.5f;

            }


            for(int column = 0;  column < WordTableColumns;column++)
            {
                table.Rows[1].Cells[column].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[column].CellFormat.Borders.Bottom.LineWidth = 1.5f;

                if(column >0)
                {
                    table.Rows[1].Cells[column].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[1].Cells[column].CellFormat.Borders.Top.LineWidth = 0.5f;
                }
                

                if (column < WordTableColumns-2)
                {
                    table.Rows[WordTableRows - 1].Cells[column].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[WordTableRows - 1].Cells[column].CellFormat.Borders.Top.LineWidth = 0.5f;
                }
                


            }

            for (int i = 3; i < (numberofgroups * 5) + 3; i = i + 5)
            {
                for (int column = 0; column < WordTableColumns; column++)
                {
                    table.Rows[i + 3].Cells[column].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[i + 3].Cells[column].CellFormat.Borders.Top.LineWidth = 0.5f;

                    table.Rows[i + 3].Cells[column].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[i + 3].Cells[column].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                }
                    

            }



        }
        public void ApplyGeneral_Groups_2Periods_Merges(IWTable table, ComparativeTable comparativeTable, int WordTableColumns, int WordTableRows)
        {
            int row = 1;
            int parametercount = comparativeTable.Parameters.Count-1;
            int testrows = parametercount / 2;

            for(int testrow = 1;testrow <= testrows; testrow++)
            {
                table.ApplyVerticalMerge(0, row, row + 8);
                

                row = row + 9;
            }


            
            int flag = 0;
            int rowctr = 1;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.IsGroup)
                {
                    continue;
                }
                table.ApplyVerticalMerge(WordTableColumns - 2, rowctr + 1, rowctr + 3);
                table.ApplyVerticalMerge(WordTableColumns - 1, rowctr + 1, rowctr + 3);

                if (rowctr != 1)
                {
                    for (int col = 1; col < WordTableColumns; col++)
                    {
                        table.Rows[rowctr].Cells[col].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[rowctr].Cells[col].CellFormat.Borders.Top.LineWidth = 0.5f;
                    }
                }

                if (flag == 0)
                {
                    rowctr = rowctr + 4;
                    flag = 1;
                }
                else if (flag == 1)
                {
                    rowctr = rowctr + 5;
                    flag = 0;
                }

                
                
                
            }

        }
        public void Insert_Inner_Header_Merges_Pathology(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns, int numberofgroups)
        {
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }


                if (parameter.NominalOrScale == "Scale")
                {
                    
                }
                else if (parameter.NominalOrScale == "Nominal")
                {


                }
            }
        }
        public void Apply_AllHeaders_Inner_Merges_Pathology(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns, int numberofgroups)
        {
            //Insert N
            AddPara_Center(table, 0, 2, "N");

            //Insert Group Name
            Parameter groupParameter = null;

            int starting_row = 3;
            int row_current_count = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                int labelinsert = starting_row;
                if(parameter.IsGroup)
                {
                    groupParameter = parameter;
                }
                else if(parameter.NominalOrScale == "Nominal")
                {
                    AddPara_NoCenter(table, starting_row, 0, parameter.Name);

                    
                    row_current_count = parameter.DIC_LablesIfNomainal.Keys.Count;
                    table.ApplyVerticalMerge(0, starting_row, starting_row + row_current_count-1);
                    table.ApplyVerticalMerge(WordTableColumns-2, starting_row, starting_row + row_current_count - 1);
                    table.ApplyVerticalMerge(WordTableColumns - 1, starting_row, starting_row + row_current_count - 1);
                    foreach (var kvp in parameter.DIC_LablesIfNomainal)
                    {
                        int key = kvp.Key;
                        string label = kvp.Value;
                        int count = parameter.ParameterValues.Count(x => x == key);

                        Addpara_NoCenterNoBOLD(table, labelinsert, 1, label);
                        LeftIntendBeforeText(table, labelinsert, 1, SetColumnWidthInCentimeters(0.3f));

                        AddPara_Center(table, labelinsert, 2, count.ToString());

                        labelinsert++;
                    }
                    for(int colctrborder = 0; colctrborder < WordTableColumns; colctrborder++)
                    {
                        if(labelinsert < WordTableRows)
                        {
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.LineWidth = 0.5f;
                        }
                        
                    }
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    AddPara_NoCenter(table, starting_row, 0, parameter.Name);
                    row_current_count = 2;
                    table.ApplyVerticalMerge(0, starting_row, starting_row + row_current_count - 1);


                    Addpara_NoCenterNoBOLD(table, labelinsert, 1, "Mean ± SD.");
                    LeftIntendBeforeText(table, labelinsert, 1, SetColumnWidthInCentimeters(0.3f));

                    Addpara_NoCenterNoBOLD(table, labelinsert+1, 1, "Median (Min. – Max.)");
                    LeftIntendBeforeText(table, labelinsert+1, 1, SetColumnWidthInCentimeters(0.3f));


                    table.ApplyVerticalMerge(WordTableColumns - 2, labelinsert, labelinsert+1);
                    table.ApplyVerticalMerge(WordTableColumns - 1, labelinsert, labelinsert + 1);


                    table.ApplyVerticalMerge(2, labelinsert, labelinsert + 1);

                    //Merging Data cells
                    int innerMergeCol = 3;
                    foreach (var kvp in parameter.GroupedParameterValues)
                    {
                        table.ApplyHorizontalMerge(labelinsert, innerMergeCol, innerMergeCol + 1);
                        table.ApplyHorizontalMerge(labelinsert+1, innerMergeCol, innerMergeCol + 1);
                        innerMergeCol += 2;

                    }


                    labelinsert += 2;

                    for (int colctrborder = 0; colctrborder < WordTableColumns; colctrborder++)
                    {
                        if (labelinsert < WordTableRows)
                        {
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.LineWidth = 0.5f;
                        }

                    }

                }







                    starting_row = starting_row + row_current_count;
            }




            AddPara_Center(table, 0, 3, groupParameter.Name);

            int col = 3;
            foreach (var kvp in groupParameter.DIC_LablesIfNomainal)
            {
                int key = kvp.Key;
                
                string label = kvp.Value;
                int count = groupParameter.ParameterValues.Count(x => x == key);

                if (count != 0)
                {
                    AddPara_Center(table, 1, col, label + "\n(n = " + count + ")");
                    AddPara_Center(table, 2, col, "No.");
                    AddPara_Center(table, 2, col + 1, "%");
                    col = col + 2;
                }
                
            }



            AddPara_Center(table, 0, WordTableColumns-1, "p");



        }


        public void Apply_AllHeaders_Inner_Merges_Pathology_IQR(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns, int numberofgroups)
        {
            //Insert N
            AddPara_Center(table, 0, 2, "N");

            //Insert Group Name
            Parameter groupParameter = null;

            int starting_row = 3;
            int row_current_count = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                int labelinsert = starting_row;
                if (parameter.IsGroup)
                {
                    groupParameter = parameter;
                }
                else if (parameter.NominalOrScale == "Nominal")
                {
                    AddPara_NoCenter(table, starting_row, 0, parameter.Name);


                    row_current_count = parameter.DIC_LablesIfNomainal.Keys.Count;
                    table.ApplyVerticalMerge(0, starting_row, starting_row + row_current_count - 1);
                    table.ApplyVerticalMerge(WordTableColumns - 2, starting_row, starting_row + row_current_count - 1);
                    table.ApplyVerticalMerge(WordTableColumns - 1, starting_row, starting_row + row_current_count - 1);
                    foreach (var kvp in parameter.DIC_LablesIfNomainal)
                    {
                        int key = kvp.Key;
                        string label = kvp.Value;
                        int count = parameter.ParameterValues.Count(x => x == key);

                        Addpara_NoCenterNoBOLD(table, labelinsert, 1, label);
                        LeftIntendBeforeText(table, labelinsert, 1, SetColumnWidthInCentimeters(0.3f));

                        AddPara_Center(table, labelinsert, 2, count.ToString());

                        labelinsert++;
                    }
                    for (int colctrborder = 0; colctrborder < WordTableColumns; colctrborder++)
                    {
                        if (labelinsert < WordTableRows)
                        {
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.LineWidth = 0.5f;
                        }

                    }
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    AddPara_NoCenter(table, starting_row, 0, parameter.Name);
                    row_current_count = 3;
                    table.ApplyVerticalMerge(0, starting_row, starting_row + row_current_count - 1);

                    Addpara_NoCenterNoBOLD(table, labelinsert , 1, "Min. – Max.");
                    LeftIntendBeforeText(table, labelinsert , 1, SetColumnWidthInCentimeters(0.3f));


                    Addpara_NoCenterNoBOLD(table, labelinsert+1, 1, "Mean ± SD.");
                    LeftIntendBeforeText(table, labelinsert+1, 1, SetColumnWidthInCentimeters(0.3f));

                    Addpara_NoCenterNoBOLD(table, labelinsert + 2, 1, "Median (IQR)");
                    LeftIntendBeforeText(table, labelinsert + 2, 1, SetColumnWidthInCentimeters(0.3f));




                    table.ApplyVerticalMerge(WordTableColumns - 2, labelinsert, labelinsert + 2);
                    table.ApplyVerticalMerge(WordTableColumns - 1, labelinsert, labelinsert + 2);


                    table.ApplyVerticalMerge(2, labelinsert, labelinsert + 2);

                    //Merging Data cells
                    int innerMergeCol = 3;
                    foreach (var kvp in parameter.GroupedParameterValues)
                    {
                        table.ApplyHorizontalMerge(labelinsert, innerMergeCol, innerMergeCol + 1);
                        table.ApplyHorizontalMerge(labelinsert + 1, innerMergeCol, innerMergeCol + 1);
                        table.ApplyHorizontalMerge(labelinsert + 2, innerMergeCol, innerMergeCol + 1);
                        innerMergeCol += 2;

                    }


                    labelinsert += 3;

                    for (int colctrborder = 0; colctrborder < WordTableColumns; colctrborder++)
                    {
                        if (labelinsert < WordTableRows)
                        {
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[labelinsert].Cells[colctrborder].CellFormat.Borders.Top.LineWidth = 0.5f;
                        }

                    }

                }







                starting_row = starting_row + row_current_count;
            }




            AddPara_Center(table, 0, 3, groupParameter.Name);

            int col = 3;
            foreach (var kvp in groupParameter.DIC_LablesIfNomainal)
            {
                int key = kvp.Key;

                string label = kvp.Value;
                int count = groupParameter.ParameterValues.Count(x => x == key);

                if (count != 0)
                {
                    AddPara_Center(table, 1, col, label + "\n(n = " + count + ")");
                    AddPara_Center(table, 2, col, "No.");
                    AddPara_Center(table, 2, col + 1, "%");
                    col = col + 2;
                }

            }



            AddPara_Center(table, 0, WordTableColumns - 1, "p");



        }
        public void ApplyPathology_Widths(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns, int numberofgroups)
        {


            for (int row = 0; row < WordTableRows; row++)
            {
                table.Rows[row].Cells[0].Width = SetColumnWidthInCentimeters(3f);
                table.Rows[row].Cells[1].Width = SetColumnWidthInCentimeters(3.65f);
                table.Rows[row].Cells[2].Width = SetColumnWidthInCentimeters(1.1f);
                table.Rows[row].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.5f);
                table.Rows[row].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.5f);

                for (int col = 3; col < WordTableColumns - 2; col++)
                {
                    table.Rows[row].Cells[col].Width = SetColumnWidthInCentimeters(1.1f);
                }

            }

        }
        public void ApplyPathology_Outer_Borders(IWTable table, ComparativeTable comparativeTable, int WordTableRows, int WordTableColumns, int numberofgroups)
        {
            for (int row = 0; row < WordTableRows; row++)
            {
                table.Rows[row].Cells[2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[row].Cells[2].CellFormat.Borders.Left.LineWidth = 1.5f;

                table.Rows[row].Cells[WordTableColumns-2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[row].Cells[WordTableColumns-2].CellFormat.Borders.Left.LineWidth = 1.5f;

            }

            for (int column = 0; column < WordTableColumns; column++)
            {
                table.Rows[2].Cells[column].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[2].Cells[column].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for (int column = 3; column < WordTableColumns-2; column++)
            {
                table.Rows[0].Cells[column].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[column].CellFormat.Borders.Bottom.LineWidth = 0.5f;

                table.Rows[1].Cells[column].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[column].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }
        }
        public void ApplyPathology_Outer_Merges(IWTable table,ComparativeTable comparativeTable,int WordTableRows ,int WordTableColumns, int numberofgroups)
        {
            //Merge empty space
            table.ApplyVerticalMerge(0, 0, 2);
            table.ApplyVerticalMerge(1, 0, 2);

            //N column
            table.ApplyVerticalMerge(2, 0, 2);


            table.ApplyHorizontalMerge(0, 3, WordTableColumns - 3);

            table.ApplyVerticalMerge(WordTableColumns-2, 0, 2);
            table.ApplyVerticalMerge(WordTableColumns -1, 0, 2);


            for(int i = 3; i < (numberofgroups*2)+3 ;i =i+2)
            {
                table.ApplyHorizontalMerge(1, i, i+1);
            }

            
        }

        public void ApplyGeneralComparativeMerges(IWTable table ,int WordTableColumns , int numberofgroups)
        {
            table.ApplyVerticalMerge(0, 0, 1);
            for(int i = 1; i<= numberofgroups*2 ;i=i+2) 
            {
                table.ApplyHorizontalMerge(0, i, i+1);
            }

            table.ApplyVerticalMerge(WordTableColumns-1 , 0, 1);
            table.ApplyVerticalMerge(WordTableColumns - 2, 0, 1);
        }
        public void ApplyGeneralPeriodsUp_Groups_ComparativeMerges(IWTable table, int WordTableColumns , int WordTableRows, int numberofgroups , ComparativeTable comparativeTable)
        {
            table.ApplyHorizontalMerge(0, 1, comparativeTable.Parameters.Count-1);

            for (int i = 3; i < (numberofgroups*5)+3; i = i+5)
            {
                table.ApplyVerticalMerge(WordTableColumns - 1, i, i+2);
                table.ApplyVerticalMerge(WordTableColumns - 2, i, i + 2);


                table.ApplyHorizontalMerge(i + 3, 1, comparativeTable.Parameters.Count - 1);

                //table.ApplyHorizontalMerge(i + 3, WordTableColumns - 2, WordTableColumns - 1);

            }

            table.ApplyVerticalMerge(WordTableColumns - 1, 0, 1);
            table.ApplyVerticalMerge(WordTableColumns - 2, 0, 1);


            //table.ApplyHorizontalMerge(WordTableRows-1, WordTableColumns - 2, WordTableColumns - 1);



        }

        public void Apply_Score_Scale_overall_merge(IWTable table , bool TS_Minmax , bool TS_MeanSD ,bool TS_Median , bool Rank , int WordTableColumns)
        {
            table.ApplyVerticalMerge(0, 0, 1);
            table.ApplyVerticalMerge(1, 0, 1);

            int TS_Count = new[] { TS_Minmax, TS_MeanSD, TS_Median }.Count(b => b);

            if (TS_Count > 1)
            {
                table.ApplyHorizontalMerge(0, 2, 1 + TS_Count);
            }


            if (Rank)
            {
                table.ApplyVerticalMerge(WordTableColumns-1, 0, 1);
            }

        }

        public void Apply_Level_Scale_overall_merge(IWTable table, int WordTableColumns)
        {
            table.ApplyVerticalMerge(0, 0, 2);

            table.ApplyHorizontalMerge(0, 1, WordTableColumns - 1);

            for (int CurrentCol = 1 ; CurrentCol < WordTableColumns; CurrentCol = CurrentCol+2)
            {
                table.ApplyHorizontalMerge(1, CurrentCol, CurrentCol+1);
            }

        }

        public void InsertHeaders_Scale_Overall(IWTable table, Tool CurrentTool, List<Tool.LevelRange> Level, int WordTableColumns)
        {
            int row = 3;

            foreach (var CurrentScale in CurrentTool.Scales)
            {
                AddPara_NoCenter(table, row, 0, CurrentScale.Scale_Full_Name);
                row++;
            }

            AddPara_Center(table, row, 0, "Overall");

            int levelctr = 0;
            for (int i = 1; i < WordTableColumns; i = i + 2)
            {
                AddPara_Center(table,1, i, Level[levelctr].Label);
                AddPara_Center(table, 2, i, "No.");
                AddPara_Center(table, 2, i+1, "%");
                levelctr++;
            }


            AddPara_Center(table, 0, 1, "Levels of " + CurrentTool.ToolName);
        }

        public void InsertGeneral_Score_Scale_overall_header_new(
    IWTable table,
    bool TS_Minmax, bool TS_MeanSD, bool TS_Median,
    bool Avg_meanSD, bool Percent_meanSD, bool Rank,
    int WordTableColumns, int minLikert, int Maxlikert)
        {
            AddPara_Center(table, 0, 1, "Score Range"); // Always at column 1

            int col = 2; // Start from column 2

            // Total Score Header (if any TS_ option is true)
            if (TS_Minmax || TS_MeanSD || TS_Median)
            {
                AddPara_Center(table, 0, col, "Total Score");

                // Row 1 details
                if (TS_Minmax)
                {
                    AddPara_Center(table, 1, col, "Min. – Max.");
                    col++;
                }
                if (TS_MeanSD)
                {
                    AddPara_Center(table, 1, col, "Mean ± SD");
                    col++;
                }
                if (TS_Median)
                {
                    AddPara_Center(table, 1, col, "Median");
                    col++;
                }
            }

            // Avg Score Header (if enabled)
            if (Avg_meanSD)
            {
                string AvgScore = $"Average Score{Convert.ToChar(11)}({minLikert} – {Maxlikert})";
                AddPara_Center(table, 0, col, AvgScore);
                AddPara_Center(table, 1, col, "Mean ± SD");
                col++;
            }

            // Percent Score Header (if enabled)
            if (Percent_meanSD)
            {
                AddPara_Center(table, 0, col, "Percent Score");
                AddPara_Center(table, 1, col, "Mean ± SD");
                col++;
            }

            // Rank Header (if enabled)
            if (Rank)
            {
                AddPara_Center(table, 0, col, "Rank");
                col++;
            }
        }

        public void InsertGeneral_Score_Scale_overall_header(IWTable table, bool TS_Minmax, bool TS_MeanSD, bool TS_Median,bool Avg_meanSD,bool Percent_meanSD ,bool Rank, int WordTableColumns , int minLikert , int Maxlikert)
        {
            AddPara_Center(table, 0, 1, "Score Range");

            bool TS = false;
            string AvgScore = "Average Score" + Convert.ToChar(11)+"("+minLikert+" - "+Maxlikert+")";

            if((TS_Minmax || TS_MeanSD || TS_Median))
            {
                TS = true;
                AddPara_Center(table, 0, 2, "Total score");
            }

            if(Avg_meanSD)
            {
                if(TS)
                {
                    if((TS_Minmax && TS_MeanSD && TS_Median))
                    {
                        AddPara_Center(table, 0, 5, AvgScore);
                        if(Percent_meanSD)
                        {
                            AddPara_Center(table, 0, 6, "Percent Score");

                        }
                    }
                    else if ((TS_Minmax && TS_MeanSD))
                    {
                        AddPara_Center(table, 0, 4, AvgScore);
                    }
                    else if ((TS_Minmax && TS_Median))
                    {
                        AddPara_Center(table, 0, 4, AvgScore);
                    }
                    else if ((TS_MeanSD && TS_Median))
                    {
                        AddPara_Center(table, 0, 4, AvgScore);
                    }
                    else
                    {
                        AddPara_Center(table, 0, 3, AvgScore);
                    }
                }
                else if(!TS)
                {
                    AddPara_Center(table, 0, 2, AvgScore);
                }
            }

        }
        public void ApplyMergesTotalScoreNoPer(IWTable table)
        {

            table.ApplyVerticalMerge(0, 0, 1);
            table.ApplyVerticalMerge(1, 0, 1);
            table.ApplyHorizontalMerge(0, 2, 4);
        }
        public void ApplyMergesTotalScorePer(IWTable table, int Arows)
        {
            
            for (int i = 2; i < Arows; i = i+6)
            {

                table.ApplyVerticalMerge(3, i, i+4);
                table.ApplyVerticalMerge(4, i, i+4);
            }
        }

        public void ApplyMergesOverallNoPer(IWTable table, int Levelcount, int WordTableRows)
        {
            for (int i = Levelcount + 2; i < WordTableRows; i++)
            {
                table.ApplyHorizontalMerge(i, 1, 2);
                

            }
        }

        public void ApplyMergesOverallPer(IWTable table , int Levelcount , int WordTableRows)
        {
            //OverallName
            table.ApplyVerticalMerge(0, 0, 1);

            //Pre Post
            table.ApplyHorizontalMerge(0, 1, 2);
            table.ApplyHorizontalMerge(0, 3, 4);

            //Test of Sig. P
            table.ApplyVerticalMerge(5, 0, 1);
            table.ApplyVerticalMerge(6, 0, 1);

            for(int i = Levelcount+2; i < WordTableRows; i++)
            {
                table.ApplyHorizontalMerge(i, 1, 2);
                table.ApplyHorizontalMerge(i, 3, 4);
                
            }
            table.ApplyVerticalMerge(5, Levelcount + 3, WordTableRows-1);
            table.ApplyVerticalMerge(6, Levelcount + 3, WordTableRows - 1);

            table.ApplyVerticalMerge(5,2, Levelcount+ 1);
            table.ApplyVerticalMerge(6, 2, Levelcount + 1);

            table.Rows[2].Cells[5].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[2].Cells[6].CellFormat.Borders.Top.LineWidth = 0.5f;

        }
        public void ApplyMerges(IWTable table , int LikertScore)
        {
            table.ApplyVerticalMerge(0, 0, 1);
            table.ApplyVerticalMerge(1, 0, 1);
            table.ApplyHorizontalMerge(0, 2, 3);
            int columns = (LikertScore * 2)+2;


            for(int i = 4; i< columns; i++)
            {
               
                table.ApplyHorizontalMerge(0, i, i+1);
                i++;
            }
            
           

            
        }
        public void DeleteColumn(List<int> columnsToRemove , IWTable table)
        {
             

            // Sort in descending order to prevent index shifting issues
            columnsToRemove.Sort();
            columnsToRemove.Reverse();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                foreach (int colIndex in columnsToRemove)
                {
                    if (table.Rows[i].Cells.Count > colIndex)
                    {
                        table.Rows[i].Cells.RemoveAt(colIndex);
                    }
                }
            }
        }
        public void Merges_Periods_Items(IWTable table , int WordTableColumns , int likertscalecount,int TestExist)
        {
            table.ApplyVerticalMerge(0, 0, 2);
            table.ApplyVerticalMerge(1, 0, 2);

            table.ApplyVerticalMerge(WordTableColumns-2, 0, 2);
            table.ApplyVerticalMerge(WordTableColumns-1, 0, 2);

            

            for(int col = 2; col < WordTableColumns-2; col = col + (likertscalecount * 2))
            {
                table.ApplyHorizontalMerge(0, col, col+ (likertscalecount*2) - 1);
            }

            for (int col = 2; col < WordTableColumns - 2; col = col + 2)
            {
                table.ApplyHorizontalMerge(1, col, col+1);
            }


        }

        public void ApplyTwoPeriodsMerges(IWTable table, int LikertScore , int numberofperiods)
        {

            int columns = (numberofperiods * LikertScore * 2) + 2;

            table.ApplyVerticalMerge(0, 0, 2);
            table.ApplyVerticalMerge(1, 0, 2);

            table.ApplyHorizontalMerge(0, 2, 2*LikertScore+1);
            table.ApplyHorizontalMerge(0, 2*LikertScore+2 , columns-1);


            

            for (int i = 2; i < columns; i++)
            {

                table.ApplyHorizontalMerge(1, i, i + 1);
                i++;
            }




        }

        public void ApplyThreePeriodsMerges(IWTable table, int LikertScore, int numberofperiods)
        {

            int columns = (numberofperiods * LikertScore * 2) + 2;

            table.ApplyVerticalMerge(0, 0, 2);
            table.ApplyVerticalMerge(1, 0, 2);

            table.ApplyHorizontalMerge(0, 2, 2 * LikertScore + 1);
            table.ApplyHorizontalMerge(0, 2 * LikertScore + 2, 4 * LikertScore + 1);
            table.ApplyHorizontalMerge(0, 4 * LikertScore + 2, columns - 1);




            for (int i = 2; i < columns; i++)
            {

                table.ApplyHorizontalMerge(1, i, i + 1);
                i++;
            }




        }

        public void ApplyBorders(IWTable table , int WordTableRows, int WordTableColumns)
        {
            for (int j = 2; j < WordTableColumns; j++)
            {
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }

            for (int k = 0; k < WordTableColumns; k++)
            {
                table.Rows[1].Cells[k].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[k].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            for (int l = 0; l < WordTableRows; l++)
            {
                table.Rows[l].Cells[1].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[l].Cells[1].CellFormat.Borders.Right.LineWidth = 1.5f;

            }
        }

        public void Borders_Periods_Items(IWTable table, int WordTableRows, int WordTableColumns)
        {
            //firstrow
            for (int j = 2; j < WordTableColumns; j++)
            {
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }

            //secondrow
            for (int j = 2; j < WordTableColumns; j++)
            {
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }

            //thirdrow
            for (int j = 0; j < WordTableColumns; j++)
            {
                table.Rows[2].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[2].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }

            //SecondCol
            for (int l = 0; l < WordTableRows; l++)
            {
                table.Rows[l].Cells[1].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[l].Cells[1].CellFormat.Borders.Right.LineWidth = 1.5f;
            }

            for (int l = 0; l < WordTableRows; l++)
            {
                table.Rows[l].Cells[WordTableColumns-2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[l].Cells[WordTableColumns-2].CellFormat.Borders.Left.LineWidth = 1.5f;
            }

        }

        public void ApplyTwoPeriodsBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int j = 2; j < WordTableColumns; j++)
            {
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;

                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }

            for (int k = 2; k < WordTableColumns; k++)
            {
                table.Rows[1].Cells[k].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[k].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }

            for (int i = 0; i < WordTableColumns; i++)
            {
                table.Rows[2].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[2].Cells[i].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }


            for (int l = 0; l < WordTableRows; l++)
            {
                table.Rows[l].Cells[1].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[l].Cells[1].CellFormat.Borders.Right.LineWidth = 1.5f;

            }
        }
        public void ApplyTotalScoreNoPBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int i = 2; i < WordTableRows; i++)
            {
                for (int j = 0; j < WordTableColumns; j++)
                {
                    table.Rows[i].Cells[j].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[i].Cells[j].CellFormat.Borders.Top.LineWidth = 0.5f;

                }

            }

            for(int l = 0;l < WordTableColumns; l++)
            {
                table.Rows[1].Cells[l].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[l].CellFormat.Borders.Top.LineWidth = 0.5f;
                
                table.Rows[1].Cells[l].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[l].CellFormat.Borders.Bottom.LineWidth = 1.5f;

            }
            for (int l = 0; l < WordTableRows; l++)
            {
                table.Rows[l].Cells[1].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[l].Cells[1].CellFormat.Borders.Left.LineWidth = 1.5f;

                

            }

        }

        public void ApplyLevel_Scale_Borders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int i = 2; i < WordTableRows; i++)
            {
                for (int j = 0; j < WordTableColumns; j++)
                {
                    table.Rows[i].Cells[j].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[i].Cells[j].CellFormat.Borders.Top.LineWidth = 0.5f;

                }

            }

            for (int l = 0; l < WordTableColumns; l++)
            {
                table.Rows[1].Cells[l].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[l].CellFormat.Borders.Top.LineWidth = 0.5f;

                table.Rows[2].Cells[l].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[2].Cells[l].CellFormat.Borders.Bottom.LineWidth = 1.5f;

            }
            for (int l = 0; l < WordTableRows; l++)
            {
                table.Rows[l].Cells[1].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[l].Cells[1].CellFormat.Borders.Left.LineWidth = 1.5f;



            }

        }
        public void ApplyTotalScorePeriodsBorders(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int i = 1; i < WordTableRows; i = i + 6)
            {
                for(int j = 0; j < WordTableColumns;j++)
                {
                    table.Rows[i].Cells[j].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[i].Cells[j].CellFormat.Borders.Top.LineWidth = 0.5f;

                    table.Rows[i+5].Cells[j].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[i+5].Cells[j].CellFormat.Borders.Top.LineWidth = 0.5f;

                }
                
            }

            
        }

        public void ApplyOverallPeriodsBorders(IWTable table, int WordTableRows, int WordTableColumns , int levelCount)
        {

            for(int i = 0; i < WordTableColumns; i++)
            {
                table.Rows[0].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;


                table.Rows[1].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;


                table.Rows[1 + levelCount].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1 + levelCount].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;

                table.Rows[5 + levelCount].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[5 + levelCount].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
            }

            






        }
        public void ApplyOverallNoPeriodsBorders(IWTable table, int WordTableRows, int WordTableColumns, int levelCount)
        {

            for (int i = 0; i < WordTableColumns; i++)
            {
                table.Rows[0].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[i].CellFormat.Borders.Bottom.LineWidth = 1.5f;


                


                table.Rows[levelCount].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[levelCount].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;

                table.Rows[4 + levelCount].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[4 + levelCount].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;


            }

            for(int i = 0;i < WordTableRows;i++)
            {
                table.Rows[i].Cells[1].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[i].Cells[1].CellFormat.Borders.Left.LineWidth = 1.5f;
            }








        }


        public void SetWidths(IWTable table , int WordTableRows , int LikertScore)
        {
            int columns = (LikertScore * 2) + 2;
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = 23;
                table.Rows[a].Cells[1].Width = 150;
                for (int i = 2 ;i < columns ;i++)
                {
                   
                    table.Rows[a].Cells[i].Width = 36;
                    
                }


            }
        }

        public void Widths_Periods_Items(IWTable table, int WordTableRows, int WordTableColumns, int likertscalecount, int TestExist , int PeriodsCount)
        {
            

            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(1f);
                table.Rows[a].Cells[1].Width = SetColumnWidthInCentimeters(5f);

                if(PeriodsCount == 2)
                {
                    for (int col = 2; col < WordTableColumns - 2; col++)
                    {
                        if(TestExist > 0)
                        {
                            if(likertscalecount == 2)
                            {
                                table.Rows[a].Cells[col].Width = SetColumnWidthInCentimeters(1f);
                                table.Rows[a].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.5f);
                                table.Rows[a].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.5f);
                            }
                            else if (likertscalecount == 3)
                            {
                                table.Rows[a].Cells[col].Width = SetColumnWidthInCentimeters(0.8f);
                                table.Rows[a].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.4f);
                                table.Rows[a].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.4f);
                            }
                        }
                        else
                        {
                            if (likertscalecount == 2)
                            {
                                table.Rows[a].Cells[col].Width = SetColumnWidthInCentimeters(1.15f);
                            }
                            else if (likertscalecount == 3)
                            {
                                table.Rows[a].Cells[col].Width = SetColumnWidthInCentimeters(0.95f);
                            }
                            
                        }
                    }

                }
                else
                {
                    for (int col = 2; col < WordTableColumns - 2; col++)
                    {
                        if (TestExist > 0)
                        {
                            table.Rows[a].Cells[col].Width = SetColumnWidthInCentimeters(0.8f);
                            table.Rows[a].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.5f);
                            table.Rows[a].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.5f);
                        }
                        else
                        {
                            table.Rows[a].Cells[col].Width = SetColumnWidthInCentimeters(0.85f);
                        }
                    }
                }
            }
        }

        public void Widths_Descriptive_Items(IWTable table, int WordTableRows, int LikertScore)
        {

            int columns = (LikertScore * 2) + 2;

            if(LikertScore > 3)
            {
                for (int a = 0; a < WordTableRows; a++)
                {
                    table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(1f);
                    table.Rows[a].Cells[1].Width = SetColumnWidthInCentimeters(5f);
                    for (int i = 2; i < columns; i++)
                    {

                        table.Rows[a].Cells[i].Width = SetColumnWidthInCentimeters(1.05f);

                    }


                }
            }
            else
            {
                for (int a = 0; a < WordTableRows; a++)
                {
                    table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(1f);
                    table.Rows[a].Cells[1].Width = SetColumnWidthInCentimeters(5f);
                    for (int i = 2; i < columns; i++)
                    {

                        table.Rows[a].Cells[i].Width = SetColumnWidthInCentimeters(1.25f);

                    }


                }
            }
            
        }

        public void SetThreePeriodsWidths(IWTable table, int WordTableRows, int LikertScore, int numberofperiods)
        {
            int columns = (numberofperiods * LikertScore * 2) + 2;
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = 23;
                table.Rows[a].Cells[1].Width = 113.4f;
                for (int i = 2; i < columns; i++)
                {

                    table.Rows[a].Cells[i].Width = 20f;
                }


            }
        }
        public void SetDescriptiveWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(3.85f);



                table.Rows[i].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.75f);
                table.Rows[i].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.75f);

                
            }



        }

        public void SetLetterWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {

            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(3.5f);

                for (int j = 1; j < WordTableColumns; j++)
                {
                    table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(3.25f);
                }
            }


        }
        public void SetMatrixCorrWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {

            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(1.75f);
                table.Rows[i].Cells[1].Width = SetColumnWidthInCentimeters(0.7f);

                for (int j = 2; j < WordTableColumns; j++)
                {
                    table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(1.65f);
                }
            }


        }
        public void SetNormalCorrWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {

            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(4.85f);
                

                for (int j = 1; j < WordTableColumns; j++)
                {
                    table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(1.75f);
                }
            }


        }
        public void SetComparative_PeriodsUp_Groups_Headers(IWTable table, int WordTableRows, int WordTableColumns, int numberofgroups, ComparativeTable comparativeTable)
        {
            if(comparativeTable.Parameters.Count-1 >= 3)
            {
                AddPara_Center(table, 0, WordTableColumns - 2, "F");
            }
            else if(comparativeTable.Parameters.Count-1 < 3)
            {
                AddPara_Center(table, 0, WordTableColumns - 2, "t");
            }

            AddPara_Center(table, 0, WordTableColumns - 1, "p");

            int col = 1;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(!parameter.IsGroup)
                {
                    AddPara_Center(table, 1, col, parameter.Name);
                    col++;
                }
            }


            int row = 2;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    var sortedKeys = parameter.DIC_LablesIfNomainal.Keys.OrderBy(key => key).ToList();
                    foreach (var key in sortedKeys)
                    {
                        int count = parameter.ParameterValues.Count(x => x == key);
                        AddPara_NoCenter(table , row , 0 , parameter.DIC_LablesIfNomainal[key] + " (n = "+count+")");

                        
                        Addpara_NoCenterNoBOLD(table, row+1, 0, "Min. – Max.");
                        LeftIntendBeforeText(table, row + 1, 0, 14.17f);

                        Addpara_NoCenterNoBOLD(table, row+2, 0, "Mean ± SD.");
                        LeftIntendBeforeText(table, row + 2, 0, 14.17f);

                        Addpara_NoCenterNoBOLD(table, row+3, 0, "Median (IQR)");
                        LeftIntendBeforeText(table, row + 3, 0, 14.17f);

                        AddPara_Center(table, row+4, 0, "Sig. bet. periods.");

                        row = row + 5;
                    }
                }
            }

            AddPara_Center(table, WordTableRows-1, 0, "t (p0)");


        }
        public void SetComparative_GroupsUp_2Periods_Widths(IWTable table, int WordTableRows, int WordTableColumns, int numberofgroups)
        {
            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(1f);
                table.Rows[i].Cells[1].Width = SetColumnWidthInCentimeters(3.75f);

                table.Rows[i].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.65f);
                table.Rows[i].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.65f);
            }



            for(int col = 2;col <= numberofgroups+1; col++)
            {
                for (int i = 0; i < WordTableRows; i++)
                {
                    table.Rows[i].Cells[col].Width = SetColumnWidthInCentimeters(3.75f);
                }
            }


        }

        public void AddHeaderComparative_GroupsUp_2Periods(IWTable table, int WordTableRows, int WordTableColumns, int numberofgroups , ComparativeTable comparativeTable)
        {
            AddPara_Center(table, 0, WordTableColumns - 2, "Test of Sig.");
            AddPara_Center(table, 0, WordTableColumns - 1, "p");
            int row = 1;
            int flag = 0;
            foreach (var paramter in comparativeTable.Parameters)
            {
                if(paramter.IsGroup)
                {
                    int col = 2;
                    foreach (var kvp in paramter.DIC_LablesIfNomainal)
                    {
                        int key = kvp.Key;
                        string label = kvp.Value;
                        int count = paramter.ParameterValues.Count(x => x == key);
                        AddPara_Center(table, 0, col, label + Convert.ToChar(11) + "(n = " + count + ")");
                        col++;
                        //WParagraph groupparagraph = table[0, col].Paragraphs[0];
                        //groupparagraph.ParagraphFormat.BeforeSpacing = 0;
                        //groupparagraph.ParagraphFormat.AfterSpacing = 0;
                    }
                }
                
                
                if (!paramter.IsGroup) 
                {
                    AddPara_NoCenter(table, row, 1, paramter.Name);

                    Addpara_NoCenterNoBOLD(table, row + 1, 1, "Min. – Max.");
                    LeftIntendBeforeText(table, row + 1, 1, 14.17f);

                    Addpara_NoCenterNoBOLD(table, row + 2, 1, "Mean ± SD.");
                    LeftIntendBeforeText(table, row + 2, 1, 14.17f);

                    Addpara_NoCenterNoBOLD(table, row + 3, 1, "Median (IQR)");
                    LeftIntendBeforeText(table, row + 3, 1, 14.17f);

                    if (flag == 0)
                    {
                        row = row + 4;
                        flag = 1;
                    }
                    else if (flag == 1)
                    {
                        row = row + 5;
                        flag = 0;
                    }
                }

            }
        }
        public void SetComparative_PeriodsUp_GroupsWidths(IWTable table, int WordTableRows, int WordTableColumns, int numberofgroups, ComparativeTable comparativeTable)
        {
            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(4f);
            }

            for(int col = 1;col <= comparativeTable.Parameters.Count-1;col++)
            {
                for (int i = 0; i < WordTableRows; i++)
                {
                    table.Rows[i].Cells[col].Width = SetColumnWidthInCentimeters(3.3f);
                }
            }

            for (int i = 0; i < WordTableRows; i++)
            {
                table.Rows[i].Cells[WordTableColumns-2].Width = SetColumnWidthInCentimeters(1.7f);
                table.Rows[i].Cells[WordTableColumns -1].Width = SetColumnWidthInCentimeters(1.7f);
            }

        }

        public void SetComparativeWidths(IWTable table, int WordTableRows, int WordTableColumns , int numberofgroups , ComparativeTable comparativeTable)
        {
            bool hasscale = false;
            bool hasnominal = false;

            foreach (var parameter in comparativeTable.Parameters)
            {
                //MessageBox.Show(parameter.Name);
                if (parameter.NominalOrScale == "Scale")
                {
                    //MessageBox.Show("Scale");
                    hasscale = true;
                    

                }
                else if(parameter.NominalOrScale == "Nominal")
                {
                    //MessageBox.Show("Nominal");
                    hasnominal = true;
                    
                }
            }


            for (int i = 0; i < WordTableRows; i++)
            {
                if(numberofgroups == 2)
                {
                    table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(4f);
                }
                else if(numberofgroups > 2)
                {
                    table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(4f);
                }
                else if(comparativeTable.HasTotalColumn)
                {
                    table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(3.75f);
                }



                if (numberofgroups == 2)
                {
                    table.Rows[i].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.75f);
                    table.Rows[i].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.75f);
                }
                else if (numberofgroups > 2)
                {
                    table.Rows[i].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.75f);
                    table.Rows[i].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.75f);
                }
                else if (comparativeTable.HasTotalColumn)
                {
                    table.Rows[i].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.75f);
                    table.Rows[i].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.75f);
                }
               
                
                if (hasscale)
                {

                    if(numberofgroups == 2)
                    {
                        for (int j = 1; j <= numberofgroups * 2; j++)
                        {
                            table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(1.75f);
                        }
                        
                    }
                    else if (numberofgroups > 2)
                    {
                        for (int j = 1; j <= numberofgroups * 2; j++)
                        {
                            table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(1.75f);
                        }
                    }
                    
                }
                else if(!hasscale)
                {
                    for (int j = 1; j <= numberofgroups * 2; j++)
                    {
                        //MessageBox.Show(hasscale.ToString());
                        table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(1.5f);

                    }
                }


                
            }
            
            

            
            
        }
        
        public void SetDescriptiveWidths(IWTable table, int WordTableRows , bool Ahasnominal)
        {
            for (int i = 0; i < WordTableRows; i++)
            {
                if(Ahasnominal)
                {
                    table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(5f);
                    table.Rows[i].Cells[1].Width = SetColumnWidthInCentimeters(2.5f);
                    table.Rows[i].Cells[2].Width = SetColumnWidthInCentimeters(2.5f);
                }
                else if(!Ahasnominal)
                {
                    table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(3.5f);
                    table.Rows[i].Cells[1].Width = SetColumnWidthInCentimeters(2.75f);
                    table.Rows[i].Cells[2].Width = SetColumnWidthInCentimeters(2.75f);
                    table.Rows[i].Cells[3].Width = SetColumnWidthInCentimeters(3.5f);
                }
                

                
            }



        }

        public void SettwoPeriodsWidths(IWTable table, int WordTableRows, int LikertScore , int numberofperiods)
        {
            int columns = (numberofperiods * LikertScore * 2) + 2;
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = 23;
                table.Rows[a].Cells[1].Width = 113.4f;
                for (int i = 2; i < columns; i++)
                {

                    table.Rows[a].Cells[i].Width = 20f;
                }


            }
        }
        public void SetTotalScoreNoPerWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            int columns = WordTableColumns;
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(5);
                table.Rows[a].Cells[1].Width = SetColumnWidthInCentimeters(2f);
                table.Rows[a].Cells[2].Width = SetColumnWidthInCentimeters(2.5f);
                table.Rows[a].Cells[3].Width = SetColumnWidthInCentimeters(2.5f);
                table.Rows[a].Cells[4].Width = SetColumnWidthInCentimeters(2.5f);
                table.Rows[a].Cells[5].Width = SetColumnWidthInCentimeters(2.5f);

            }
        }
        public void Score_Overall_Widths(IWTable table, int WordTableRows, int WordTableColumns , bool Rank)
        {
            
            
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(4f);
                table.Rows[a].Cells[1].Width = SetColumnWidthInCentimeters(2f);
                if(Rank)
                {
                    table.Rows[a].Cells[WordTableColumns-1].Width = SetColumnWidthInCentimeters(1.4f);
                }

            }

            for (int i = 0; i < WordTableRows; i++)
            {
                for (int columnctr = 2; columnctr < WordTableColumns - 1; columnctr++)
                {
                    if(WordTableColumns < 8)
                    {
                        table.Rows[i].Cells[columnctr].Width = SetColumnWidthInCentimeters(2.75f);
                    }
                    else
                    {
                        table.Rows[i].Cells[columnctr].Width = SetColumnWidthInCentimeters(2.5f);
                    }
                    
                }
            }
           
        }

        public void Level_Overall_Widths(IWTable table, int WordTableRows, int WordTableColumns)
        {


            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(5f);
                
                

            }

            for (int i = 0; i < WordTableRows; i++)
            {
                for (int columnctr = 1; columnctr < WordTableColumns; columnctr++)
                {
                    table.Rows[i].Cells[columnctr].Width = SetColumnWidthInCentimeters(1.5f);

                }
            }

        }

        public void SetTotalScorePerWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            int columns = WordTableColumns;
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = 154f;          
                table.Rows[a].Cells[1].Width = 84f;
                table.Rows[a].Cells[2].Width = 84f;
                table.Rows[a].Cells[3].Width = 49f;
                table.Rows[a].Cells[4].Width = 49f;

            }
        }
        public float SetColumnWidthInCentimeters(float widthInCentimeters)
        {
            // Convert centimeters to points
            float widthInPoints = widthInCentimeters * 28.3465f;
            return widthInPoints;
        }
        public void SetOverallPerWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            int columns = WordTableColumns;
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(5f);
                table.Rows[a].Cells[1].Width = SetColumnWidthInCentimeters(1.5f);
                table.Rows[a].Cells[2].Width = SetColumnWidthInCentimeters(1.5f);
                table.Rows[a].Cells[3].Width = SetColumnWidthInCentimeters(1.5f);
                table.Rows[a].Cells[4].Width = SetColumnWidthInCentimeters(1.5f);
                table.Rows[a].Cells[5].Width = SetColumnWidthInCentimeters(1.75f);
                table.Rows[a].Cells[6].Width = SetColumnWidthInCentimeters(1.75f);

            }
        }
        public void SetOverallNoPerWidths(IWTable table, int WordTableRows, int WordTableColumns)
        {
            
            for (int a = 0; a < WordTableRows; a++)
            {
                table.Rows[a].Cells[0].Width = SetColumnWidthInCentimeters(6f);
                table.Rows[a].Cells[1].Width = SetColumnWidthInCentimeters(2f);
                table.Rows[a].Cells[2].Width = SetColumnWidthInCentimeters(2f);
                
            }
        }
        public void AddHorizontalTitlesOverallPer(IWTable table)
        {
            AddPara_Center(table, 0, 1 , "Pre");
            AddPara_Center(table, 0, 3, "Post");
            InsertHighlightTestName(table, 0, 5, "Test Of Sig.");
            AddPara_Center(table, 0, 6, "p");

            AddPara_Center(table, 1, 1, "No.");
            AddPara_Center(table, 1, 3, "No.");
            AddPara_Center(table, 1, 2, "%");
            AddPara_Center(table, 1, 4, "%");
        }

        public void AddHorizontalTitlesOverallNoPer(IWTable table)
        {
            
            

            AddPara_Center(table, 0, 1, "No.");
            AddPara_Center(table, 0, 2, "%");
        }

        public void AddVerticalTitlesOverallPer(IWTable table , int LevelCount , int TotalScorefrom , int TotalScoreto , int AvgScoreFrom , int AvgScoreto )
        {
            int LastColumnLevel = LevelCount + 1;
            AddPara_NoCenter(table, LastColumnLevel+1, 0, "Total Score (" + TotalScorefrom + " – " + TotalScoreto + ")");
            LeftIntendBeforeText(table, LastColumnLevel+1, 0, SetColumnWidthInCentimeters(0.5f));


            Addpara_NoCenterNoBOLD(table, LastColumnLevel + 2, 0, "Min. – Max.");
            Addpara_NoCenterNoBOLD(table, LastColumnLevel + 3, 0, "Mean ± SD.");
            Addpara_NoCenterNoBOLD(table, LastColumnLevel + 4, 0, "Median");
            LeftIntendBeforeText(table, LastColumnLevel + 2, 0, 28.35f);
            LeftIntendBeforeText(table, LastColumnLevel + 3, 0, 28.35f);
            LeftIntendBeforeText(table, LastColumnLevel + 4, 0, 28.35f);

            AddPara_NoCenter(table, LastColumnLevel + 5, 0, "Average Score (" + AvgScoreFrom + " – " + AvgScoreto + ") \r\n(Mean ± SD.)");
            LeftIntendBeforeText(table, LastColumnLevel + 5, 0, 14.17f);
        }
        public void AddVerticalTitlesOverallNoPer(IWTable table, int LevelCount, int TotalScorefrom, int TotalScoreto, int AvgScoreFrom, int AvgScoreto)
        {
            int LastColumnLevel = LevelCount + 1;
            AddPara_NoCenter(table, LastColumnLevel , 0, "Total Score (" + TotalScorefrom + " – " + TotalScoreto + ")");
            //LeftIntendBeforeText(table, LastColumnLevel , 0, SetColumnWidthInCentimeters(0.5f));


            Addpara_NoCenterNoBOLD(table, LastColumnLevel + 1, 0, "Min. – Max.");
            Addpara_NoCenterNoBOLD(table, LastColumnLevel + 2, 0, "Mean ± SD.");
            Addpara_NoCenterNoBOLD(table, LastColumnLevel + 3, 0, "Median");
            LeftIntendBeforeText(table, LastColumnLevel + 1, 0, SetColumnWidthInCentimeters(0.5f));
            LeftIntendBeforeText(table, LastColumnLevel + 2, 0, SetColumnWidthInCentimeters(0.5f));
            LeftIntendBeforeText(table, LastColumnLevel + 3, 0, SetColumnWidthInCentimeters(0.5f));

            AddPara_NoCenter(table, LastColumnLevel + 4, 0, "Average Score (" + AvgScoreFrom + " – " + AvgScoreto + ") \r\n(Mean ± SD.)");
            
        }
        public void LeftIntendBeforeText(IWTable table , int row , int column , float intendvalue)
        {
            WTableCell cell = table.Rows[row].Cells[column];
            WParagraph paragraph = cell.Paragraphs[0] as WParagraph;
            paragraph.ParagraphFormat.LeftIndent = intendvalue;
        }

        public void InsertScoreRangesTotalScore(IWTable table, int WordTableRows , string ScoreRange)
        {
            
            
            AddPara_Center(table, WordTableRows, 1, ScoreRange);
            
            
        }
        public void ConstStringsTotalScorePeriods(IWTable table, int WordTableRows, int WordTableColumns , int TotalScorefrom , int TotalScoreto , int AvgScoreFrom , int AvgScoreto , int itemindex)
        {
            int TotRow = 2 + (itemindex * 6);
            int TotRowMinMax = 2 + (itemindex * 6)+1;
            int TotRowMeanStd = 2 + (itemindex * 6) + 2;
            int TotRowMedian = 2 + (itemindex * 6) + 3;
            int AvgRowMeanStd = 2 + (itemindex * 6) + 4;

            


            
            AddPara_NoCenter(table, TotRow, 0, "Total Score (" + TotalScorefrom +" – "+TotalScoreto +")");
            LeftIntendBeforeText(table, TotRow, 0, 14.17f);
            
            
            Addpara_NoCenterNoBOLD(table, TotRowMinMax, 0, "Min. – Max.");
            Addpara_NoCenterNoBOLD(table, TotRowMeanStd, 0, "Mean ± SD.");
            Addpara_NoCenterNoBOLD(table, TotRowMedian, 0, "Median");
            LeftIntendBeforeText(table, TotRowMinMax, 0, 28.35f);
            LeftIntendBeforeText(table, TotRowMeanStd, 0, 28.35f);
            LeftIntendBeforeText(table, TotRowMedian, 0, 28.35f);

            AddPara_NoCenter(table, AvgRowMeanStd, 0, "Average Score ("+AvgScoreFrom+" – "+AvgScoreto+") \r\n(Mean ± SD.)");
            LeftIntendBeforeText(table, AvgRowMeanStd, 0, 14.17f);
            



        }
        List<string> HeadersList = new List<string>();
        
        public void Add_Headers_Center(IWTable table , string CurrentDomainName , int LikertScore)
        {
            int headerlistctr = 0;
            HeadersList.Add("Strongly Disagree");
            HeadersList.Add("Disagree");
            HeadersList.Add("Not sure");
            HeadersList.Add("Agree");
            HeadersList.Add("Strongly Agree");
            HeadersList.Add("6");
            HeadersList.Add("7");


            //table[0, 0].AddParagraph().AppendText("Q").CharacterFormat.Bold = true;
            //table[0, 1].AddParagraph().AppendText(CurrentDomainName).CharacterFormat.Bold = true;

            AddPara_Center(table, 0, 0, "Q");
            AddPara_Center(table, 0, 1, CurrentDomainName);



            for (int i = 2;i <= LikertScore*2;i = i+2) 
            {
                
                string headerlisttext = HeadersList[headerlistctr];

                AddPara_Center(table, 0, i, headerlisttext);
                
                headerlistctr++;
            }    
                
         
            
        }

        public void Add_Header_Items_Nusring(IWTable table, int LikertScore , string ScaleName , Tool CurrentTool)
        {
            
            


            AddPara_Center(table, 0, 0, "Q");
            AddPara_Center(table, 0, 1, ScaleName);


            int Dictionatryctr = 0;
            var likertKeys = CurrentTool.LikertScale.Keys.ToList();
            for (int i = 2; i <= LikertScore * 2; i = i + 2)
            {

                string headerlisttext = CurrentTool.LikertScale[likertKeys[Dictionatryctr]];

                AddPara_Center(table, 0, i, headerlisttext);

                Dictionatryctr++;
            }



        }

        public void Header_Periods_Items(IWTable table, int LikertScore, Tool.Scale CurrentScale, Tool CurrentTool , int WordTableColumns , int itemcount)
        {
            AddPara_Center(table, 0, 0, "Q");
            AddPara_Center(table, 0, 1, CurrentScale.Scale_Full_Name);


            //Period Titles
            List<string> PeriodNames = new List<string>();

            foreach (var CurrentToolPeriod in CurrentTool.PeriodsTools)
            {
                PeriodNames.Add(CurrentToolPeriod.ToolName);
            }

            int periodnamectr = 0;
            for(int col = 2; col < WordTableColumns-2; col = col + (LikertScore*2))
            {
                AddPara_Center(table, 0, col, PeriodNames[periodnamectr]);
                periodnamectr++;
            }



            //Likert Names
            List<string> likertValues = CurrentTool.LikertScale.Values.ToList(); // Store dictionary values in a list
            int likertCount = likertValues.Count; // Get count for cycling

            List<string> resultList = new List<string>();

            for (int i = 2, j = 0; i < WordTableColumns - 2; i += 2, j++)
            {
                //mod resets the list
                AddPara_Center(table, 1, i, likertValues[j % likertCount]);
            }


            //Add No %
            for (int col = 2; col < WordTableColumns - 2; col++)
            {
                AddPara_Center(table, 2, col, "No.");
                col++;
                AddPara_Center(table, 2, col, "%");
            }

            // Add Question count
            int rowQcount = 3;
            int QuestionCount = 1;
            for(int currentitem = 0;currentitem < itemcount; currentitem++)
            {
                AddPara_Center(table, rowQcount, 0, QuestionCount.ToString());
                rowQcount++;
                QuestionCount++;
            }

        }
        public void Add_GeneralHeaders_Descriptive(IWTable table, bool hasnominal)
        {
            if(hasnominal)
            {
                AddPara_Center(table, 0, 1, "No.");
                AddPara_Center(table, 0, 2, "%");
            }
            else if(!hasnominal)
            {
                AddPara_Center(table, 0, 1, "Min. – Max.");
                AddPara_Center(table, 0, 2, "Mean ± SD.");
                AddPara_Center(table, 0, 3, "Median (IQR)");
            }
        }

        public void Add_GeneralHeaders_Comparative_Center(IWTable table , int WordTableRows , int WordTableColumns , int numberofgroups)
        {
            //InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test Of Sig.");
            AddPara_Center(table, 0, WordTableColumns - 1, "p");
            for (int i = 1; i <= numberofgroups*2; i++)
            {
                if(i%2 != 0)
                {
                    AddPara_Center(table, 1, i , "No.");

                }
                else if(i%2 == 0)
                {
                    AddPara_Center(table, 1, i, "%");
                }
            }


        }
        public void Add_Headers_TwoPeriods_Center(IWTable table, string CurrentDomainName, int Columns ,  int LikertScore)
        {
            int headerlistctr = 0;
            HeadersList.Add("Strongly Disagree");
            HeadersList.Add("Disagree");
            HeadersList.Add("Not sure");
            HeadersList.Add("Agree");
            HeadersList.Add("Strongly Agree");
            HeadersList.Add("Strongly Disagree");
            HeadersList.Add("Disagree");
            HeadersList.Add("Not sure");
            HeadersList.Add("Agree");
            HeadersList.Add("Strongly Agree");


            //table[0, 0].AddParagraph().AppendText("Q").CharacterFormat.Bold = true;
            //table[0, 1].AddParagraph().AppendText(CurrentDomainName).CharacterFormat.Bold = true;

            AddPara_Center(table, 0, 0, "Q");
            AddPara_Center(table, 0, 1, CurrentDomainName);

            AddPara_Center(table, 0, 2, "Pre");
            AddPara_Center(table, 0, 2*LikertScore+2, "Post");

            for (int i = 2; i <= Columns-1; i = i + 2)
            {

                string headerlisttext = HeadersList[headerlistctr];

                AddPara_Center(table, 1, i, headerlisttext);

                headerlistctr++;
            }



        }
        public void Add_Headers_ThreePeriods_Center(IWTable table, string CurrentDomainName, int Columns, int LikertScore)
        {
            int headerlistctr = 0;
            HeadersList.Add("Strongly Disagree");
            HeadersList.Add("Disagree");
            HeadersList.Add("Not sure");
            HeadersList.Add("Agree");
            HeadersList.Add("Strongly Agree");
            HeadersList.Add("Strongly Disagree");
            HeadersList.Add("Disagree");
            HeadersList.Add("Not sure");
            HeadersList.Add("Agree");
            HeadersList.Add("Strongly Agree");
            HeadersList.Add("Strongly Disagree");
            HeadersList.Add("Disagree");
            HeadersList.Add("Not sure");
            HeadersList.Add("Agree");
            HeadersList.Add("Strongly Agree");


            //table[0, 0].AddParagraph().AppendText("Q").CharacterFormat.Bold = true;
            //table[0, 1].AddParagraph().AppendText(CurrentDomainName).CharacterFormat.Bold = true;

            AddPara_Center(table, 0, 0, "Q");
            AddPara_Center(table, 0, 1, CurrentDomainName);

            AddPara_Center(table, 0, 2, "Pre");
            AddPara_Center(table, 0, 2 * LikertScore + 2, "Post");
            AddPara_Center(table, 0, 4 * LikertScore + 2, "3M");

            for (int i = 2; i < Columns - 2; i = i + 2)
            {

                string headerlisttext = HeadersList[headerlistctr];

                AddPara_Center(table, 1, i, headerlisttext);

                headerlistctr++;
            }



        }



        public void AddNo_perc_Center(IWTable table , int WordTableColumns)
        {
            for (int f = 2; f < WordTableColumns; f++)
            {
                if (f % 2 == 0)
                {
                    AddPara_Center(table , 1 , f , "No.");
                }
                else if (f % 2 != 0)
                {
                    AddPara_Center(table, 1, f, "%");
                }
                
            }
        }

        public void AddNo_perc_Center_periods(IWTable table, int WordTableColumns)
        {
            for (int f = 2; f < WordTableColumns-2; f++)
            {
                if (f % 2 == 0)
                {
                    AddPara_Center(table, 2, f, "No.");
                }
                else if (f % 2 != 0)
                {
                    AddPara_Center(table,2, f, "%");
                }

            }
        }
        

        public void FormatTable(IWTable table, float fontSize)
        {
            
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    // Check if the cell has any content
                    if (cell.Paragraphs.Count == 0)
                    {
                        // If the cell is empty, create a new paragraph and text range
                        WParagraph paragraph = (WParagraph)cell.AddParagraph();
                        paragraph.ParagraphFormat.BeforeSpacing = 4;
                        paragraph.ParagraphFormat.AfterSpacing = 2;

                        // Create a new text range in the paragraph
                        WTextRange textRange = (WTextRange)paragraph.AppendText("");
                        paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                        // Set the font properties for the text range
                        textRange.CharacterFormat.FontName = "Times New Roman";
                        textRange.CharacterFormat.FontSize = fontSize;
                    }
                    else
                    {
                        // Iterate items in cell and set horizontal alignment
                        foreach (WParagraph paragraph in cell.Paragraphs)
                        {
                            paragraph.ParagraphFormat.BeforeSpacing = 4;
                            paragraph.ParagraphFormat.AfterSpacing = 2;

                            foreach (ParagraphItem item in paragraph.ChildEntities)
                            {
                                if (item is WTextRange)
                                {
                                    WTextRange text = item as WTextRange;
                                    text.CharacterFormat.FontName = "Times New Roman";
                                    text.CharacterFormat.FontSize = fontSize;
                                }
                            }
                        }
                    }
                }
            }
        }

        public void FormatTable_Periods_Items(IWTable table,int PeriodsCount, int likertscalecount,int TestExist)
        {
            if (PeriodsCount == 2)
            {
                if (TestExist > 0)
                {
                    if (likertscalecount == 2)
                    {
                        FormatTableCustom(table, 11, 4, 2);
                        LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    }
                    else if (likertscalecount == 3)
                    {
                        FormatTableCustom(table, 10.5f, 4, 2);
                        LeftAndRightCellMarginCustom(table, 0f, 0f);

                    }
                }
                else
                {
                    if (likertscalecount == 2)
                    {
                        FormatTableCustom(table, 11, 4, 2);
                        LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    }
                    else if (likertscalecount == 3)
                    {
                        FormatTableCustom(table, 11, 4, 2);
                        LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    }

                }

            }
            else
            {
                if (TestExist > 0)
                {
                    FormatTableCustom(table, 10, 4, 2);
                    LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);

                }
                else
                {
                    FormatTableCustom(table, 10.5f, 4, 2);
                    LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                }
            }


            
        }
        public void FormatTableCustom(IWTable table, float fontSize , int beforeSpacing , int AfterSpacing)
        {
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    // Check if the cell has any content
                    if (cell.Paragraphs.Count == 0)
                    {
                        // If the cell is empty, create a new paragraph and text range
                        WParagraph paragraph = (WParagraph)cell.AddParagraph();


                        WTextRange textRange = (WTextRange)paragraph.AppendText(" ");
                        paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                        paragraph.ParagraphFormat.BeforeSpacing = beforeSpacing;
                        paragraph.ParagraphFormat.AfterSpacing = AfterSpacing;

                        // Create a new text range in the paragraph
                        

                        // Set the font properties for the text range
                        textRange.CharacterFormat.FontName = "Times New Roman";
                        textRange.CharacterFormat.FontSize = fontSize;
                    }
                    else
                    {
                        // Iterate items in cell and set horizontal alignment
                        foreach (WParagraph paragraph in cell.Paragraphs)
                        {
                            paragraph.ParagraphFormat.BeforeSpacing = beforeSpacing;
                            paragraph.ParagraphFormat.AfterSpacing = AfterSpacing;

                            foreach (ParagraphItem item in paragraph.ChildEntities)
                            {
                                if (item is WTextRange)
                                {
                                    WTextRange text = item as WTextRange;
                                    text.CharacterFormat.FontName = "Times New Roman";
                                    text.CharacterFormat.FontSize = fontSize;
                                }
                            }
                        }
                    }
                }
            }
        }


        public void CorrMatrixCustom(IWTable table, int WordTableRows, int WordTableColumns, ComparativeTable comparativeTable)
        {
            int parametercount = comparativeTable.Parameters.Count;
            int ParameterNameRows = 1;
            int ParameterNameColumns = 2;

            foreach (Parameter parameter in comparativeTable.Parameters)
            {
                //adding parameter name on left side (column)
                AddPara_NoCenter(table, ParameterNameRows, 0, parameter.Name);

                //making it 0 and 0 to intend the name
                LeftAndRightCellMarginCustom(table, 0, 0);

                //merging cells
                table.ApplyVerticalMerge(0, ParameterNameRows, ParameterNameRows + 1);

                AddPara_Center(table, ParameterNameRows, 1, "r");


                if (comparativeTable.CorreType == "Spearman")
                {
                    SubSuperScriptText(table, ParameterNameRows, 1, Color.Empty, "s", "Sub");
                }
                AddPara_Center(table, ParameterNameRows + 1, 1, "p");

                AddPara_Center(table, 0, ParameterNameColumns, parameter.Name);

                ParameterNameRows += 2;
                ParameterNameColumns++;
            }

        }
        public void CorrNormalCustom(IWTable table , int WordTableRows , int WordTableColumns , ComparativeTable comparativeTable)
        {
            int DependentParameters = 0;
            int IndependParameters = 0;

            foreach (Parameter parameter in comparativeTable.Parameters)
            {
                if(parameter.IsIndependentCorr)
                {
                    IndependParameters++;
                }
                else if(parameter.NormalOrAbnormal == "Normal" || parameter.NormalOrAbnormal == "Abnormal")
                {
                    DependentParameters++;
                }
            }

            int InsertColumn = 1;
            int InsertRow = 2;  
            foreach (Parameter parameter in comparativeTable.Parameters)
            {
                //adding parameter name on left side (column)

                if (parameter.NormalOrAbnormal == "Normal" || parameter.NormalOrAbnormal == "Abnormal")
                {
                    AddPara_Center(table, 0, InsertColumn, parameter.Name);

                    if(parameter.NormalOrAbnormal == "Normal")
                    {
                        AddPara_Center(table, 1, InsertColumn, "r");
                        AddPara_Center(table, 1, InsertColumn+1, "p");
                    }
                    else if(parameter.NormalOrAbnormal == "Abnormal")
                    {
                        AddPara_Center(table, 1, InsertColumn, "r");
                        SubSuperScriptText(table, 1, InsertColumn, Color.Empty, "s", "Sub");

                        AddPara_Center(table, 1, InsertColumn + 1, "p");
                    }

                    InsertColumn = InsertColumn + 2;
                }
                
                if(parameter.IsIndependentCorr)
                {
                    AddPara_NoCenter(table, InsertRow, 0, parameter.Name);
                    InsertRow++;
                }
               

            }

        }
        public void Font(IWTable table , float fontsize)
        {
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row
                foreach (WTableCell cell in row.Cells)
                {
                    //Iterate items in cell and set horizontal alignment
                    foreach (WParagraph paragraph in cell.Paragraphs)
                    {
                        paragraph.ParagraphFormat.BeforeSpacing = 4;
                        paragraph.ParagraphFormat.AfterSpacing = 2;
                        foreach (ParagraphItem item in paragraph.ChildEntities)
                        {
                            if (item is WTextRange)
                            {
                                WTextRange text = item as WTextRange;
                                text.CharacterFormat.FontName = "Times New Roman";
                                text.CharacterFormat.FontSize = fontsize;

                                
                                
                            }
                        }
                    }
                }
            }
        }

        public void LeftAndRightCellMarginCustom(IWTable table , float left , float right)
        {
            WTable tableProperties = table as WTable;

            tableProperties.TableFormat.Paddings.Left = left;
            tableProperties.TableFormat.Paddings.Right = right;



        }

        public void LeftAndRightCellMargin(IWTable table)
        {
            WTable tableProperties = table as WTable;

            tableProperties.TableFormat.Paddings.Left = 0;
            tableProperties.TableFormat.Paddings.Right = 0;



        }
        public void HighlightCellContent(IWTable table, int row, int WordTableColumns)
        {
            WParagraph firstparagraph = null;
            try
            {
                firstparagraph = table[row, 0].Paragraphs[0];
            }
            catch(Exception)
            {
                firstparagraph = null;
            }
           

            if (firstparagraph != null)
            {

                // Retrieve the text content of the cell
                string cellText = firstparagraph.Text;

                firstparagraph.Text = "";


                WTextRange newTextRange = (WTextRange)firstparagraph.AppendText(cellText);
                newTextRange.CharacterFormat.HighlightColor = Color.Yellow;
            }

            // Get the first paragraph of the specified cell
            for (int i = 1; i < WordTableColumns-2; i = i+2)
            {

                if (table[row, i].Paragraphs.Count > 0)
                {
                    WParagraph paragraph = table[row, i].Paragraphs[0];

                    if (paragraph != null)
                    {

                        // Retrieve the text content of the cell
                        string cellText = paragraph.Text;

                        paragraph.Text = "";


                        WTextRange newTextRange = (WTextRange)paragraph.AppendText(cellText);
                        newTextRange.CharacterFormat.HighlightColor = Color.Yellow;
                    }
                }
                

                    

                
                
            }
            
        }


        public void HighlightCellContent_Paper(IWTable table, int row, int WordTableColumns)
        {

            WParagraph firstparagraph = table[row, 0].Paragraphs[0];

            if (firstparagraph != null)
            {

                // Retrieve the text content of the cell
                string cellText = firstparagraph.Text;

                firstparagraph.Text = "";


                WTextRange newTextRange = (WTextRange)firstparagraph.AppendText(cellText);
                newTextRange.CharacterFormat.HighlightColor = Color.Yellow;
            }

            // Get the first paragraph of the specified cell
            for (int i = 1; i < WordTableColumns - 2; i ++)
            {

                if (table[row, i].Paragraphs.Count > 0)
                {
                    WParagraph paragraph = table[row, i].Paragraphs[0];

                    if (paragraph != null)
                    {

                        // Retrieve the text content of the cell
                        string cellText = paragraph.Text;

                        paragraph.Text = "";


                        WTextRange newTextRange = (WTextRange)paragraph.AppendText(cellText);
                        newTextRange.CharacterFormat.HighlightColor = Color.Yellow;
                    }
                }






            }

        }
        public void InsertHighlightTestName(IWTable table , int InsertRow , int InsertColumn , string TestName)
        {
            WParagraph testparaHighlight = (WParagraph)table[InsertRow, InsertColumn].AddParagraph();
            WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestName);
            HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;
            HighlightTest.CharacterFormat.Bold = true;
            testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            table[InsertRow, InsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
        }

        public void InsertHighlightTestName_Nobold(IWTable table, int InsertRow, int InsertColumn, string TestName)
        {
            WParagraph testparaHighlight = (WParagraph)table[InsertRow, InsertColumn].AddParagraph();

            testparaHighlight.ParagraphFormat.BeforeSpacing = 0;
            testparaHighlight.ParagraphFormat.AfterSpacing = 0;
            
            WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestName);
            HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;
            HighlightTest.CharacterFormat.Bold = false;

           
            
            testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            table[InsertRow, InsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
        }

        public void SubSuperScriptText(IWTable table , int insertrow , int insertcolumn , Color Highlightcolor , string SubSuperScripttext , string SuperOrSubscript)
        {
            WParagraph paragraph = (WParagraph)table[insertrow, insertcolumn].Paragraphs[0];
            WTextRange asteriskt = (WTextRange)paragraph.AppendText(SubSuperScripttext);
            if(SuperOrSubscript == "Super")
            {
                asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
            }
            else if(SuperOrSubscript == "Sub")
            {
                asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;
            }
            
            asteriskt.CharacterFormat.HighlightColor = Highlightcolor;
        }

        public void InsertPairwiseF(IWTable table , int insertrow , int insertcolumn , string text , WParagraph paragraph , bool Isbold ,bool Iscenter, Color HighlightColor , Color TextColor)
        {
            
           
            WTextRange p = (WTextRange)paragraph.AppendText(text);
            p.CharacterFormat.HighlightColor = HighlightColor;
            p.CharacterFormat.Bold = Isbold;
            p.CharacterFormat.TextColor = TextColor;

            if (Iscenter)
            {
                paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[insertrow, insertcolumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            }
            else if (!Iscenter)
            {
                table[insertrow, insertcolumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            }

        }

        public void InsertTest_P_Letters(IWTable table, int TestInsertrow, int TestInsertColumn, string[] TestValueFrm)
        {
            double testValue = 0.0;

            //MessageBox.Show(TestValueFrm[1]);
            if (TestValueFrm[1] == "<0.001")
            {

                WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;

                WTextRange asteriskt = (WTextRange)testparaHighlight.AppendText("*");
                asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                asteriskt.CharacterFormat.HighlightColor = Color.Yellow;

                testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                WParagraph testpara = (WParagraph)table[TestInsertrow+1, TestInsertColumn ].AddParagraph();
                testpara.AppendText("<0.001");
                WTextRange asteriskp = (WTextRange)testpara.AppendText("*");
                asteriskp.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

                testpara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[TestInsertrow+1, TestInsertColumn ].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            }




            if (double.TryParse(TestValueFrm[1], out testValue))
            {
                if (testValue < 0.05)
                {
                    // Code for when TestValueFrm[1] is less than 0.05
                    WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                    WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                    HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;

                    WTextRange asteriskt = (WTextRange)testparaHighlight.AppendText("*");
                    asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    asteriskt.CharacterFormat.HighlightColor = Color.Yellow;

                    testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                    WParagraph testpara = (WParagraph)table[TestInsertrow+1, TestInsertColumn ].AddParagraph();
                    testpara.AppendText(TestValueFrm[1]);
                    WTextRange asteriskp = (WTextRange)testpara.AppendText("*");
                    asteriskp.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

                    testpara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow+1, TestInsertColumn ].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                }
                else
                {
                    WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                    WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                    HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;
                    testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                    Addpara_CenterNoBOLD(table, TestInsertrow+1, TestInsertColumn , TestValueFrm[1]);

                }

            }

        }
        public void InsertTest_P(IWTable table , int TestInsertrow , int TestInsertColumn,string[] TestValueFrm)
        {
            double testValue = 0.0;

            //MessageBox.Show(TestValueFrm[1]);
            if (TestValueFrm[1] == "<0.001")
            {

                WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;

                WTextRange asteriskt = (WTextRange)testparaHighlight.AppendText("*");
                asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                asteriskt.CharacterFormat.HighlightColor = Color.Yellow;



                testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                WParagraph testpara = (WParagraph)table[TestInsertrow, TestInsertColumn + 1].AddParagraph();
                testpara.AppendText("<0.001");
                WTextRange asteriskp = (WTextRange)testpara.AppendText("*");
                asteriskp.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

                testpara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[TestInsertrow, TestInsertColumn + 1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            }


            

            if (double.TryParse(TestValueFrm[1], out testValue))
            {
                if (testValue < 0.05)
                {
                    // Code for when TestValueFrm[1] is less than 0.05
                    WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                    WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                    HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;

                    WTextRange asteriskt = (WTextRange)testparaHighlight.AppendText("*");
                    asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    asteriskt.CharacterFormat.HighlightColor = Color.Yellow;

                    testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                    WParagraph testpara = (WParagraph)table[TestInsertrow, TestInsertColumn + 1].AddParagraph();
                    testpara.AppendText(TestValueFrm[1]);
                    WTextRange asteriskp = (WTextRange)testpara.AppendText("*");
                    asteriskp.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

                    testpara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow, TestInsertColumn + 1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                }
                else
                {
                    WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                    WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                    HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;
                    testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                    Addpara_CenterNoBOLD(table, TestInsertrow, TestInsertColumn + 1, TestValueFrm[1]);

                }

            }

        }

        public void InsertTest_P_Corr(IWTable table, int TestInsertrow, int TestInsertColumn, string[] TestValueFrm)
        {
            double testValue = 0.0;

            //MessageBox.Show(TestValueFrm[1]);
            if (TestValueFrm[1] == "<0.001")
            {

                WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                //HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;

                WTextRange asteriskt = (WTextRange)testparaHighlight.AppendText("*");
                asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                //asteriskt.CharacterFormat.HighlightColor = Color.Yellow;

                testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                WParagraph testpara = (WParagraph)table[TestInsertrow+1, TestInsertColumn].AddParagraph();
                testpara.AppendText("<0.001");
                WTextRange asteriskp = (WTextRange)testpara.AppendText("*");
                asteriskp.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

                testpara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[TestInsertrow+1, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            }




            if (double.TryParse(TestValueFrm[1], out testValue))
            {
                if (testValue < 0.05)
                {
                    // Code for when TestValueFrm[1] is less than 0.05
                    WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                    WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                    HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;

                    WTextRange asteriskt = (WTextRange)testparaHighlight.AppendText("*");
                    asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    asteriskt.CharacterFormat.HighlightColor = Color.Yellow;

                    testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                    WParagraph testpara = (WParagraph)table[TestInsertrow+1, TestInsertColumn].AddParagraph();
                    testpara.AppendText(TestValueFrm[1]);
                    WTextRange asteriskp = (WTextRange)testpara.AppendText("*");
                    asteriskp.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

                    testpara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow+1, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                }
                else
                {
                    WParagraph testparaHighlight = (WParagraph)table[TestInsertrow, TestInsertColumn].AddParagraph();
                    WTextRange HighlightTest = (WTextRange)testparaHighlight.AppendText(TestValueFrm[0]);
                    HighlightTest.CharacterFormat.HighlightColor = Color.Yellow;
                    testparaHighlight.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                    table[TestInsertrow, TestInsertColumn].CellFormat.VerticalAlignment = VerticalAlignment.Middle;


                    Addpara_CenterNoBOLD(table, TestInsertrow+1, TestInsertColumn, TestValueFrm[1]);

                }

            }

        }


        public void AddQuestionNo(IWTable table , int AQuestioncount)
        {
            int QuestionNum = AQuestioncount;
            //MessageBox.Show(QuestionNum.ToString());
            int Question_index = 1;
            for (int row = 2; row <= QuestionNum + 1; row++)
            {
                AddPara_Center(table, row, 0, Question_index.ToString());
                
                Question_index++;
            }
        }
        public void AddQuestionNo_Periods(IWTable table, int AQuestioncount)
        {
            int QuestionNum = AQuestioncount;
            //MessageBox.Show(QuestionNum.ToString());
            int Question_index = 1;
            for (int row = 3; row <= QuestionNum + 2; row++)
            {
                AddPara_Center(table, row, 0, Question_index.ToString());

                Question_index++;
            }
        }

        public void AddParaCombined(IWTable table , int WordTableRows, int WordTableColumns, string Text, bool Isbold, bool Iscenter , Color HighlightColor , Color TextColor)
        {
            WParagraph paragraph = (WParagraph)table[WordTableRows, WordTableColumns].AddParagraph();
            
            WTextRange textRange = (WTextRange)paragraph.AppendText(Text);

            textRange.CharacterFormat.HighlightColor = HighlightColor;
            textRange.CharacterFormat.Bold = Isbold;
            textRange.CharacterFormat.TextColor = TextColor;

            if(Iscenter)
            {
                paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                table[WordTableRows, WordTableColumns].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            }
            else if(!Iscenter) 
            {
                table[WordTableRows, WordTableColumns].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            }
            
        }
        public void AddPara_Center(IWTable table , int WordTableRows , int WordTableColumns , string Text)
        {
            table[WordTableRows, WordTableColumns].AddParagraph().AppendText(Text).CharacterFormat.Bold = true;
            WParagraph paragraph = table[WordTableRows, WordTableColumns].Paragraphs[0];
            paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            table[WordTableRows, WordTableColumns].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
        }

        

        public void AddPara_NoCenter(IWTable table, int WordTableRows, int WordTableColumns, string Text)
        {
            table[WordTableRows, WordTableColumns].AddParagraph().AppendText(Text).CharacterFormat.Bold = true;
            WParagraph paragraph = table[WordTableRows, WordTableColumns].Paragraphs[0];
            
            table[WordTableRows, WordTableColumns].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
        }

        public void Addpara_CenterNoBOLD(IWTable table, int WordTableRows, int WordTableColumns, string Text)
        {
            table[WordTableRows, WordTableColumns].AddParagraph().AppendText(Text);
            WParagraph paragraph = table[WordTableRows, WordTableColumns].Paragraphs[0];
            paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            table[WordTableRows, WordTableColumns].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
        }

        public void Addpara_NoCenterNoBOLD(IWTable table, int WordTableRows, int WordTableColumns, string Text)
        {
            table[WordTableRows, WordTableColumns].AddParagraph().AppendText(Text);
            WParagraph paragraph = table[WordTableRows, WordTableColumns].Paragraphs[0];
            
            table[WordTableRows, WordTableColumns].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
        }


        

        public string SaveWord()
        {
            try
            {

                

                // Initialize the SaveFileDialog
                saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Word Documents (*.docx)|*.docx"; // Filter for Word documents
                saveFileDialog1.DefaultExt = "docx"; // Default file extension
                saveFileDialog1.AddExtension = true; // Automatically add extension if the user omits it
                string filePath = "";
                // Show the dialog and check if the user clicked OK
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    filePath = saveFileDialog1.FileName; // Get the full file path chosen by the user

                    using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        // Save the document to the file stream
                        document.Save(fileStream, FormatType.Docx);
                        MessageBox.Show("Tables Created!");
                    }
                }

                return filePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }


        }

    }
}
