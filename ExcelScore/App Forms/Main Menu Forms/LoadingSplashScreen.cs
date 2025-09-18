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
using System.Threading;
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
                pic_SpssPathVerify.Image = Resources.check1;
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
                pic_ExcelPathVerify.Image = Resources.check1;
            }
            else
            {
            }
        }
        

        private void btn_Settings_Click(object sender, EventArgs e)
        {
           
        }

        private void btn_Exit_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }
        public bool timerdone = false;
        

        private void pic_verify_Click(object sender, EventArgs e)
        {
            CheckPCID();
            if (AcceptedPC)
            {
                LoadExcel test = new LoadExcel();
                test.ShowDialog();
            }
            else
            {
                Application.Exit();
            }
        }

        private void pic_SpssPathCopy_Click(object sender, EventArgs e)
        {
            if(SpssReaderStat.spssFilePath !=null)
            {
                Clipboard.SetText(SpssReaderStat.spssFilePath);
                pic_SpssPathCopy.Image = Resources.CopySuccess;

            }
        }

        private void pic_ExcelPathCopy_Click(object sender, EventArgs e)
        {
            if(ExcelFunctions.filepath != null)
            {
                Clipboard.SetText(ExcelFunctions.filepath);
                pic_SpssPathCopy.Image = Resources.CopySuccess;
            }
        }

        private void OpenInExplorer(string path)
        {
            if (System.IO.File.Exists(path))
            {
                // Open Explorer with the file selected
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\"");
            }
            else if (System.IO.Directory.Exists(path))
            {
                // If it's a folder path, just open it
                System.Diagnostics.Process.Start("explorer.exe", path);
            }
            else
            {
                MessageBox.Show("Path not found: " + path, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pic_SpssPathOpen_Click(object sender, EventArgs e)
        {
            if (SpssReaderStat.spssFilePath != null)
            {
                OpenInExplorer(SpssReaderStat.spssFilePath);
            }
            
        }

        private void pic_ExcelPathOpen_Click(object sender, EventArgs e)
        {
            if(ExcelFunctions.filepath != null)
            {
                OpenInExplorer(ExcelFunctions.filepath);
            }
        }
    }
}
