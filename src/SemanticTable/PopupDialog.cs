using System;
using System.Drawing;
using System.Windows.Forms;

namespace SemanticTable
{
    internal sealed class PopupDialog : Form
    {
        internal PopupDialog()
        {
            SuspendLayout();
            AutoScaleDimensions = new SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.Sizable;
        }

        protected override void OnLoad(EventArgs e)
        {
            ResumeLayout(true);
            PerformAutoScale();
            base.OnLoad(e);
            var area = Screen.FromControl(this).WorkingArea;
            var scale = Math.Max(DeviceDpi / 96F, Font.Height / 13F);
            MinimumSize = new Size(Math.Min(area.Width, (int)(360 * scale)),
                Math.Min(area.Height, (int)(200 * scale)));
            Size = new Size(Math.Min(Width, area.Width), Math.Min(Height, area.Height));
            Location = new Point(area.Left + (area.Width - Width) / 2,
                area.Top + (area.Height - Height) / 2);
        }
    }
}
