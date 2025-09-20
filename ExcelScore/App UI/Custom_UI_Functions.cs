using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.App_UI
{
    public class Custom_UI_Functions
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);
        public static void Make_Panel_Draggable(Control control, Form form)
        {
            control.MouseDown += (sender, e) =>
            {
                ReleaseCapture();
                SendMessage(form.Handle, 0x112, 0xf012, 0);
            };
        }
    }
}
