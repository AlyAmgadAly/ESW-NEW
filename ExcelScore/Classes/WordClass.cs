using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Humanizer;

using Aspose.Cells;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.CompilerServices;
using CenterSpace.NMath.Core;

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

        public void ApplyGeneralComparativeMerges(IWTable table ,int WordTableColumns , int numberofgroups)
        {
            table.ApplyVerticalMerge(0, 0, 1);
            //table.ApplyHorizontalMerge(0, 1, 2);
            //table.ApplyHorizontalMerge(0, 3, 4);
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
                    table.Rows[i].Cells[0].Width = SetColumnWidthInCentimeters(3.75f);
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
                    table.Rows[i].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.65f);
                    table.Rows[i].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.65f);
                }
                else if (comparativeTable.HasTotalColumn)
                {
                    table.Rows[i].Cells[WordTableColumns - 1].Width = SetColumnWidthInCentimeters(1.65f);
                    table.Rows[i].Cells[WordTableColumns - 2].Width = SetColumnWidthInCentimeters(1.65f);
                }
               
                
                if (hasscale)
                {

                    if(numberofgroups == 2)
                    {
                        for (int j = 1; j <= numberofgroups * 2; j++)
                        {
                            table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(1.65f);
                        }
                        
                    }
                    else if (numberofgroups > 2)
                    {
                        for (int j = 1; j <= numberofgroups * 2; j++)
                        {
                            table.Rows[i].Cells[j].Width = SetColumnWidthInCentimeters(1.5f);
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

            for (int i = 2; i <= Columns - 1; i = i + 2)
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
            for (int f = 2; f < WordTableColumns; f++)
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
                        paragraph.ParagraphFormat.BeforeSpacing = beforeSpacing;
                        paragraph.ParagraphFormat.AfterSpacing = AfterSpacing;

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


        

        public void SaveWord()
        {
            try
            {
                // Initialize the SaveFileDialog
                saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Word Documents (*.docx)|*.docx"; // Filter for Word documents
                saveFileDialog1.DefaultExt = "docx"; // Default file extension
                saveFileDialog1.AddExtension = true; // Automatically add extension if the user omits it

                // Show the dialog and check if the user clicked OK
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog1.FileName; // Get the full file path chosen by the user

                    using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        // Save the document to the file stream
                        document.Save(fileStream, FormatType.Docx);
                        MessageBox.Show("Tables Created!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

    }
}
