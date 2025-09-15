using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Custom_UI
{
    public class FormUI
    {
        /// <summary>
        /// Apply rounded corners to a form.
        /// </summary>
        public static void ApplyRoundedCorners(Form form, int radius)
        {
            if (form == null) return;

            GraphicsPath path = new GraphicsPath();
            path.StartFigure();

            // Top-left arc
            path.AddArc(new Rectangle(0, 0, radius, radius), 180, 90);
            // Top edge
            path.AddLine(radius, 0, form.Width - radius, 0);
            // Top-right arc
            path.AddArc(new Rectangle(form.Width - radius, 0, radius, radius), 270, 90);
            // Right edge
            path.AddLine(form.Width, radius, form.Width, form.Height - radius);
            // Bottom-right arc
            path.AddArc(new Rectangle(form.Width - radius, form.Height - radius, radius, radius), 0, 90);
            // Bottom edge
            path.AddLine(form.Width - radius, form.Height, radius, form.Height);
            // Bottom-left arc
            path.AddArc(new Rectangle(0, form.Height - radius, radius, radius), 90, 90);
            // Left edge
            path.AddLine(0, form.Height - radius, 0, radius);

            path.CloseFigure();
            form.Region = new Region(path);
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
                dragging = true;
                dragCursor = Cursor.Position;
                dragForm = form.Location;
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
