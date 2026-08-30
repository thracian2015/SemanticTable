using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Windows.Forms;

namespace SemanticTable
{
    public sealed class AddIn : IExcelAddIn
    {
        public void AutoOpen() { }
        public void AutoClose() => FieldsRibbon.ClosePane();
    }

    [ComVisible(true)]
    public sealed class FieldsRibbon : ExcelRibbon
    {
        private static CustomTaskPane _pane;

        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs><tab id='ctfTab' label='Semantic Table'>
    <group id='ctfGroup' label='Semantic Table'>
      <button id='ctfOpen' label='Fields' size='large' getImage='GetSemanticTableImage' onAction='OpenFields'/>
      <button id='ctfSettings' label='Settings' size='large' getImage='GetSettingsImage' onAction='ShowSettings'/>
      <button id='ctfAbout' label='About' onAction='ShowAbout'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";

        public void OpenFields(IRibbonControl control)
        {
            try
            {
                if (_pane == null)
                {
                    _pane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(FieldsPane), "Semantic Table Fields");
                    _pane.Width = 360;
                }
                _pane.Visible = true;
                ((FieldsPane)_pane.ContentControl).AttachToSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Semantic Table failed while creating or showing the task pane.\r\n\r\n" + ex.Message,
                    "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ShowAbout(IRibbonControl control)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                ?? assembly.GetName().Version?.ToString(3);
            if (!string.IsNullOrEmpty(version)) version = version.Split('+')[0];
            MessageBox.Show(
                "Semantic Table\r\nVersion " + version +
                "\r\n\r\nBuild and filter regular Excel connected tables from Power BI semantic models." +
                "\r\n\r\nCopyright (c) 2026 Prologika, LLC" +
                "\r\nLicensed under the MIT License." +
                "\r\n\r\nThis software is provided as-is, without warranty of any kind, express or implied." +
                "\r\nSee LICENSE and THIRD-PARTY-NOTICES.md for details." +
                "\r\n\r\nGitHub: https://github.com/thracian2015/SemanticTable",
                "About Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ShowSettings(IRibbonControl control)
        {
            try
            {
                if (_pane == null)
                {
                    _pane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(FieldsPane), "Semantic Table Fields");
                    _pane.Width = 360;
                    _pane.Visible = true;
                    ((FieldsPane)_pane.ContentControl).AttachToSelection();
                }
                ((FieldsPane)_pane.ContentControl).ShowSettingsDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Semantic Table could not open Settings.\r\n\r\n" + ex.Message,
                    "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public object GetSettingsImage(IRibbonControl control)
        {
            var image = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.FromArgb(70, 90, 110), 3f))
            using (var brush = new SolidBrush(Color.FromArgb(70, 90, 110)))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.DrawEllipse(pen, 8, 8, 16, 16);
                g.FillEllipse(brush, 13, 13, 6, 6);
                for (var i = 0; i < 8; i++)
                {
                    var angle = i * Math.PI / 4;
                    var x1 = 16 + (float)Math.Cos(angle) * 10;
                    var y1 = 16 + (float)Math.Sin(angle) * 10;
                    var x2 = 16 + (float)Math.Cos(angle) * 14;
                    var y2 = 16 + (float)Math.Sin(angle) * 14;
                    g.DrawLine(pen, x1, y1, x2, y2);
                }
            }
            return image;
        }

        public object GetSemanticTableImage(IRibbonControl control)
        {
            var image = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(image))
            using (var gridPen = new Pen(Color.FromArgb(16, 124, 65), 2f))
            using (var linkPen = new Pen(Color.FromArgb(45, 125, 210), 2f))
            using (var nodeBrush = new SolidBrush(Color.FromArgb(45, 125, 210)))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.DrawRoundedRectangle(gridPen, new Rectangle(2, 5, 20, 23), 3);
                g.DrawLine(gridPen, 2, 12, 22, 12);
                g.DrawLine(gridPen, 2, 19, 22, 19);
                g.DrawLine(gridPen, 9, 5, 9, 28);
                g.DrawLine(gridPen, 16, 5, 16, 28);
                g.DrawLine(linkPen, 22, 10, 27, 7);
                g.DrawLine(linkPen, 22, 16, 28, 16);
                g.DrawLine(linkPen, 22, 22, 27, 25);
                g.FillEllipse(nodeBrush, 25, 4, 6, 6);
                g.FillEllipse(nodeBrush, 26, 13, 6, 6);
                g.FillEllipse(nodeBrush, 25, 22, 6, 6);
            }
            return image;
        }

        internal static void ClosePane()
        {
            if (_pane == null) return;
            _pane.Delete();
            _pane = null;
        }
    }

    internal static class GraphicsExtensions
    {
        public static void DrawRoundedRectangle(this Graphics graphics, Pen pen, Rectangle bounds, int radius)
        {
            var diameter = radius * 2;
            using (var path = new GraphicsPath())
            {
                path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
                path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
                path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                graphics.DrawPath(pen, path);
            }
        }
    }
}
