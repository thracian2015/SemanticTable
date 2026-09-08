using System;
using System.Drawing;
using System.Windows.Forms;

namespace SemanticTable
{
    internal sealed class ConnectionDialog : Form
    {
        private readonly TextBox _connection;
        internal string ConnectionString => _connection.Text;

        internal ConnectionDialog(string current, bool creatingTable)
        {
            SuspendLayout();
            AutoScaleDimensions = new SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            Text = creatingTable ? "Create Semantic Table connection" : "Semantic Table connection";
            ClientSize = new Size(800, 300);
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 3
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                AutoSize = true, Dock = DockStyle.Fill,
                Text = creatingTable
                    ? "Enter the complete MSOLAP connection string for the semantic model. Replace <workspace> and <semantic-model> with actual values."
                    : "Edit the complete connection string used by Semantic Table. Replace the <workspace> placeholder with the Power BI workspace name. This setting is saved for this table."
            }, 0, 0);
            _connection = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true, Text = current ?? ""
            };
            layout.Controls.Add(_connection, 0, 1);
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft
            };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            var ok = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            layout.Controls.Add(buttons, 0, 2);
            Controls.Add(layout);
            AcceptButton = ok;
            CancelButton = cancel;
            ResumeLayout(true);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            var area = Screen.FromControl(this).WorkingArea;
            MinimumSize = new Size(Math.Min(area.Width, (int)(420 * CurrentAutoScaleDimensions.Width / 6F)),
                Math.Min(area.Height, (int)(240 * CurrentAutoScaleDimensions.Height / 13F)));
            Size = new Size(Math.Min(Width, area.Width), Math.Min(Height, area.Height));
            Location = new Point(area.Left + (area.Width - Width) / 2, area.Top + (area.Height - Height) / 2);
        }
    }
}
