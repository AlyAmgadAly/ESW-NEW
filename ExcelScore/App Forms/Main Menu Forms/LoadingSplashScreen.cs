using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Math;
using ExcelScore.Properties;
using ExcelScore.StatClasses;
using Google.OrTools.LinearSolver;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using System.Xml.Linq;

namespace ExcelScore.App_Forms.Main_Menu_Forms
{
    public partial class LoadingSplashScreen : Form
    {
        public LoadingSplashScreen()
        {
            InitializeComponent();
        }
        List<string> MyPCs = new List<string>
        {

            "50026B73817770D6",
            "50026B7381195A47",
            "            9VYLNL4E",
            "5636355932474844202020202020202020202020",
            "30533556444a5157303634393034202020202020",
            "            Z3T5S767" ,
            "S1LJJ1NQC12283" ,
            "202020202020202020202020563542564a535850",
            "     WD-WCAV3E296961",
            "     WD-WCAP9C351984" ,
            "            Z3TXT3DA",
            "            ZA428LY9",
            "     WD-WCAV9L3N7RFP",
            "4C530201731113100112",
            "      S1VCJ90Z621250",
            "S1LJJDWQ602716",
            "            W9AT0ETT",
            "S0MRJDSP802594"
        };

        private static ManagementObjectSearcher baseboardSearcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
        private static ManagementObjectSearcher motherboardSearcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
        public bool AcceptedPC = false;
        public void CheckPCID()
        {
            string currentPCID = SerialNumber();
            // MessageBox.Show(currentPCID);
            if (MyPCs.Contains(currentPCID))
            {

                AcceptedPC = true;
            }
            else
            {
                AcceptedPC = false;
            }
        }

        public string SerialNumber()
        {
            try
            {
                foreach (ManagementObject queryObj in motherboardSearcher.Get())
                {
                    return queryObj["SerialNumber"].ToString();

                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        //        Background(#222831) → calming and focus-friendly.

        //Secondary (#393E46) → adds subtle layering, keeps things from looking flat.

        //Accent (#00ADB5, teal) → perfect balance: teal = both blue (focus) and green (balance/refresh).

        //Text (#EEEEEE) → high contrast, readable, not tiring like pure white (#FFFFFF).
        private void LoadingSplashScreen_Load(object sender, EventArgs e)
        {

        }
        SpssReaderStat spssReaderStat = new SpssReaderStat();
        
        private void btn_SPSSImport_Click(object sender, EventArgs e)
        {
            spssReaderStat.GetSPSS_Path_Parameters();

            if (SpssReaderStat.spssFilePath != null)
            {
                lbl_SpssPath.Text = ShortenPath(SpssReaderStat.spssFilePath);
            }
            else
            {
                lbl_SpssPath.Text = "NA";
            }
        }

        private void btn_SpssToExcel_Click(object sender, EventArgs e)
        {
            spssReaderStat.Convert_Spss_to_Excel();
        }

        private void btn_ExcelImport_Click(object sender, EventArgs e)
        {
            ExcelFunctions.ImportExcelFile();

            if (ExcelFunctions.filepath != null)
            {
                lbl_ExcelPath.Text = ShortenPath(ExcelFunctions.filepath);
            }
            else
            {
                lbl_ExcelPath.Text = "NA";
            }
        }
        public static string ShortenPath(string fullPath, int maxLength = 50)
        {
            if (string.IsNullOrEmpty(fullPath)) return fullPath;

            if (fullPath.Length <= maxLength)
                return fullPath;

            string root = System.IO.Path.GetPathRoot(fullPath) ?? "";
            string fileName = System.IO.Path.GetFileName(fullPath);

            // Remaining middle part
            string middle = fullPath.Substring(root.Length, fullPath.Length - root.Length - fileName.Length);

            // If still too long, replace with "..."
            if ((root + "...\\" + fileName).Length > maxLength)
                return root + "...\\" + fileName;

            return root + "...\\" + fileName;
        }

        private void btn_Settings_Click(object sender, EventArgs e)
        {
           
        }

        private void btn_Exit_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }
        public bool timerdone = false;
        private void timer_progress_Tick(object sender, EventArgs e)
        {
            pnl_progress.Width += 40;
            if (pnl_progress.Width >= 799)
            {
                timer_progress.Stop();
                timerdone = true;
                CheckPCID();
                if(AcceptedPC)
                {
                    pic_verify.Image =  Resources.check;
                    lbl_verify.Text = "Verified";
                    lbl_verify.ForeColor = Color.Green;

                    pnl_progress.BackColor = Color.Green;
                }
                else
                {
                    Application.Exit();
                }
            }

        }
    }
}
