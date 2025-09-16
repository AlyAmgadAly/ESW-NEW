using Humanizer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class SplashScreen : Form
    {
        public SplashScreen()
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
            //ayhaga test
        };    
        

        public bool timerdone = false;
        private void timer_progress_Tick(object sender, EventArgs e)
        {
            pnl_progress.Width += 40;
            if (pnl_progress.Width >= 799)
            {
                timer_progress.Stop();
                timerdone = true;
                CheckPCID();
            }
            

        }
        private static ManagementObjectSearcher baseboardSearcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
        private static ManagementObjectSearcher motherboardSearcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
        private static string getCpuID()
        {
            ManagementClass management = new ManagementClass("win32_processor");
            ManagementObjectCollection managementObjectCollection = management.GetInstances();

            foreach (var managementObject in managementObjectCollection)
            {
                var cpuid = managementObject.Properties["processorID"].Value.ToString();
                return cpuid;
            }

            return "";
        }
        public void CheckPCID()
        {
            string currentPCID = SerialNumber();          
           // MessageBox.Show(currentPCID);
            if(MyPCs.Contains(currentPCID))
            {

                LoadExcel send = new LoadExcel();
                send.Show();
                this.Hide();
            }
            else
            {
                Application.Exit();
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
        private void SplashScreen_Load(object sender, EventArgs e)
        {
            
        }
    }
}
