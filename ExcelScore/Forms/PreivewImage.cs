using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class PreivewImage : Form
    {
        public Image SelectedImage { get; set; }
        public PreivewImage()
        {
            InitializeComponent();
        }

        private void PreivewImage_Load(object sender, EventArgs e)
        {
            pic_ImageToPreview.Image = SelectedImage;
            pic_ImageToPreview.Size = this.ClientSize;
        }

        private void pic_ImageToPreview_Click(object sender, EventArgs e)
        {
            
        }
    }
}
