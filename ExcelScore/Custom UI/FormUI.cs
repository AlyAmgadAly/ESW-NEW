using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ExcelScore.Custom_UI
{
    public class FormUI
    {
        [DllImport("dwmapi.dll")]
        private static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS pMargins);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        [DllImport("dwmapi.dll")]
        private static extern int DwmIsCompositionEnabled(ref int pfEnabled);

        [StructLayout(LayoutKind.Sequential)]
        private struct MARGINS
        {
            public int leftWidth;
            public int rightWidth;
            public int topHeight;
            public int bottomHeight;
        }

        /// <summary>
        /// Apply drop shadow to a borderless form.
        /// </summary>
        public static void ApplyShadow(Form form)
        {
            int enabled = 0;
            DwmIsCompositionEnabled(ref enabled);

            if (enabled == 1)
            {
                int val = 2; // Enable shadow
                DwmSetWindowAttribute(form.Handle, 2, ref val, sizeof(int));

                MARGINS margins = new MARGINS()
                {
                    leftWidth = 1,
                    rightWidth = 1,
                    topHeight = 1,
                    bottomHeight = 1
                };

                DwmExtendFrameIntoClientArea(form.Handle, ref margins);
            }
        }


        /// <summary>
        /// Apply rounded corners to a form with anti-aliased edges.
        /// </summary>
        public static void ApplyRoundedCorners(Form form, int radius)
        {
            if (form == null || form.Width <= 0 || form.Height <= 0) return;

            form.Region = new Region(BuildRoundedPath(form.Width, form.Height, radius));

            form.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var brush = new SolidBrush(form.BackColor))
                {
                    using (GraphicsPath paintPath = BuildRoundedPath(form.Width, form.Height, radius))
                    {
                        e.Graphics.FillPath(brush, paintPath);
                    }
                }
            };
        }
        private static GraphicsPath BuildRoundedPath(int width, int height, int radius)
        {
            GraphicsPath path = new GraphicsPath();

            path.StartFigure();
            path.AddArc(new Rectangle(0, 0, radius, radius), 180, 90);                           // Top-left
            path.AddLine(radius, 0, width - radius, 0);                                         // Top edge
            path.AddArc(new Rectangle(width - radius, 0, radius, radius), 270, 90);             // Top-right
            path.AddLine(width, radius, width, height - radius);                                // Right edge
            path.AddArc(new Rectangle(width - radius, height - radius, radius, radius), 0, 90); // Bottom-right
            path.AddLine(width - radius, height, radius, height);                               // Bottom edge
            path.AddArc(new Rectangle(0, height - radius, radius, radius), 90, 90);             // Bottom-left
            path.AddLine(0, height - radius, 0, radius);                                        // Left edge
            path.CloseFigure();

            return path;
        }

        /// <summary>
        /// Make a control (like a panel) draggable, so it can be used as a custom title bar.
        /// </summary>
        public static void EnableDrag(Control control, Form form)
        {
            bool dragging = false;
            Point dragCursor = Point.Empty;
            Point dragForm = Point.Empty;

            control.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    dragging = true;
                    dragCursor = Cursor.Position;
                    dragForm = form.Location;
                }
            };

            control.MouseMove += (s, e) =>
            {
                if (dragging)
                {
                    Point diff = Point.Subtract(Cursor.Position, new Size(dragCursor));
                    form.Location = Point.Add(dragForm, new Size(diff));
                }
            };

            control.MouseUp += (s, e) => { dragging = false; };
        }
    }
}
