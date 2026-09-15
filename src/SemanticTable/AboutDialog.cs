using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace SemanticTable
{
    internal static class AboutDialog
    {
        internal static Form Create(string version)
        {
            var dialog = new PopupDialog
            {
                Text = "About Semantic Table", ClientSize = new Size(510, 420),
                MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false,
                StartPosition = FormStartPosition.CenterParent
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, Padding = new Padding(20), ColumnCount = 1,
                RowCount = 9, AutoScroll = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var i = 0; i < 8; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Action<string, int> addText = (text, space) => layout.Controls.Add(new Label
            {
                Text = text, AutoSize = true, Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, space)
            });
            addText("Semantic Table\r\nVersion " + version, 16);
            addText("Build and filter regular Excel connected tables from Power BI semantic models.", 16);
            addText("Copyright (c) 2026 Prologika, LLC", 0);
            layout.Controls.Add(CreateLink("https://prologika.com", "https://prologika.com", dialog));
            addText("Licensed under the MIT License.", 16);
            addText("This software is provided as-is, without warranty of any kind, express or implied.\r\nSee LICENSE and THIRD-PARTY-NOTICES.md for details.", 16);
            layout.Controls.Add(CreateLink("GitHub: https://github.com/thracian2015/SemanticTable",
                "https://github.com/thracian2015/SemanticTable", dialog));
            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK,
                AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(0, 16, 0, 0) };
            layout.Controls.Add(ok);
            dialog.AcceptButton = ok;
            dialog.CancelButton = ok;
            dialog.Controls.Add(layout);
            dialog.Shown += (_, __) =>
            {
                var preferred = layout.GetPreferredSize(new Size(layout.ClientSize.Width, 0));
                var area = Screen.FromControl(dialog).WorkingArea;
                dialog.ClientSize = new Size(dialog.ClientSize.Width,
                    Math.Min(preferred.Height, area.Height - (dialog.Height - dialog.ClientSize.Height)));
                dialog.Top = area.Top + (area.Height - dialog.Height) / 2;
            };
            return dialog;
        }

        private static LinkLabel CreateLink(string text, string url, Form owner)
        {
            var link = new LinkLabel { Text = text, AutoSize = true, Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 4), TabStop = true };
            link.Links.Clear();
            link.Links.Add(text.IndexOf(url, StringComparison.Ordinal), url.Length, url);
            link.LinkClicked += (_, e) =>
            {
                try
                {
                    Process.Start(new ProcessStartInfo((string)e.Link.LinkData) { UseShellExecute = true });
                    e.Link.Visited = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(owner, "Could not open the link.\r\n" + ex.Message,
                        "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            return link;
        }
    }
}

