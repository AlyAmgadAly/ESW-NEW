using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ExcelScore.Forms
{
    public partial class CombineDomains : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public List<NewDomainClass> AllDomainsCombined = new List<NewDomainClass>();
        public CombineDomains()
        {
            InitializeComponent();
        }
        
        private void CombineDomains_Load(object sender, EventArgs e)
        {
            
            foreach(var domain in AllDomainsCombined) 
            {
                listbox_Alldomains.Items.Add(domain.DomainName);
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(listbox_Alldomains.Items.Count>0)
            {
                string SelectedItem = listbox_Alldomains.GetItemText(listbox_Alldomains.SelectedItem);
                listbox_Alldomains.Items.Remove(listbox_Alldomains.SelectedItems[0]);
                foreach(var domain in AllDomainsCombined.ToList())
                {
                    if(SelectedItem == domain.DomainName) 
                    {
                        AllDomainsCombined.Remove(domain);
                    }
                }
            }
            
        }

        private void CombineDomains_FormClosed(object sender, FormClosedEventArgs e)
        {
            NewAddDomain newAddDomain = new NewAddDomain();
            newAddDomain.AllDomains = AllDomainsCombined;
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
    }
}
