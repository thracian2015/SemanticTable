using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SemanticTable
{
    internal sealed class FilterRow : UserControl
    {
        private readonly FieldFilter _filter;
        private readonly ComboBox _mode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 90 };
        private readonly ComboBox _operator = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 105 };
        private readonly Control _value1;
        private readonly Control _value2;
        private readonly Button _selectValues = new Button { Left = 104, Top = 35, Width = 249, Height = 24 };
        private readonly Func<SemanticField, bool, string, IReadOnlyList<string>> _loadValues;
        private readonly Action _changed;
        private readonly Action _stateChanged;
        private bool _updatingValues;

        public FilterRow(FieldFilter filter, Action<FieldFilter> remove,
            Func<SemanticField, bool, string, IReadOnlyList<string>> loadValues, Action changed, Action stateChanged)
        {
            _filter = filter;
            _loadValues = loadValues;
            _changed = changed;
            _stateChanged = stateChanged;
            if (_filter.Mode != "Advanced") _filter.Mode = "Basic";
            if (_filter.Mode == "Basic") _filter.Operator = "Equals";
            if (_filter.Operator == "Before") _filter.Operator = "Less Than";
            if (_filter.Operator == "After") _filter.Operator = "Greater Than";
            if (_filter.Values == null) _filter.Values = new List<string>();
            if (_filter.Values.Count == 0 && !string.IsNullOrWhiteSpace(_filter.Value)) _filter.Values.Add(_filter.Value);
            Height = 100;
            Width = 365;
            Margin = Padding.Empty;
            BorderStyle = BorderStyle.FixedSingle;

            var title = new Label
            {
                Text = filter.Field.Table + "  >  " + filter.Field.Name,
                Left = 8, Top = 7, Width = 275, AutoEllipsis = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            var clear = new Button
            {
                Left = 295, Top = 3, Width = 30, Height = 25, FlatStyle = FlatStyle.Flat,
                Image = CreateEraserIcon(), TabStop = false, AccessibleName = "Clear filter selection",
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            clear.FlatAppearance.BorderSize = 0;
            var delete = new Button
            {
                Text = "×", Left = 325, Top = 3, Width = 30, Height = 25,
                FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            delete.FlatAppearance.BorderSize = 0;
            var toolTip = new ToolTip();
            toolTip.SetToolTip(clear, "Clear filter selection");
            clear.Click += (_, __) => ClearSelection();
            delete.Click += (_, __) => remove(filter);

            _mode.Left = 8;
            _mode.Top = 35;
            _mode.Items.AddRange(new object[] { "Basic", "Advanced" });
            _mode.SelectedItem = _filter.Mode;

            _operator.Left = 104;
            _operator.Top = 35;
            _operator.Items.AddRange(Operators(filter.Field).Cast<object>().ToArray());
            _operator.SelectedItem = _operator.Items.Contains(filter.Operator) ? filter.Operator : "Equals";
            if (_operator.SelectedIndex < 0) _operator.SelectedIndex = 0;

            _value1 = CreateValueControl(filter.Value, 215, 35);
            _value2 = CreateValueControl(filter.Value2, 215, 65);
            _mode.SelectedIndexChanged += (_, __) =>
            {
                var previousCondition = DaxQueryBuilder.ConditionSignature(_filter);
                _filter.Mode = Convert.ToString(_mode.SelectedItem);
                if (_filter.Mode == "Basic") _filter.Operator = "Equals";
                UpdateValueControls();
                if (!string.Equals(previousCondition, DaxQueryBuilder.ConditionSignature(_filter), StringComparison.Ordinal))
                    _changed?.Invoke();
                else
                    _stateChanged?.Invoke();
            };
            _operator.SelectedIndexChanged += (_, __) =>
            {
                var previousCondition = DaxQueryBuilder.ConditionSignature(_filter);
                _filter.Operator = Convert.ToString(_operator.SelectedItem);
                UpdateValueControls();
                if (!string.Equals(previousCondition, DaxQueryBuilder.ConditionSignature(_filter), StringComparison.Ordinal))
                    _changed?.Invoke();
                else
                    _stateChanged?.Invoke();
            };
            _selectValues.Click += (_, __) => ShowValuePicker();
            _selectValues.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _value1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _value2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            UpdateValueControls();

            Controls.Add(title);
            Controls.Add(clear);
            Controls.Add(delete);
            Controls.Add(_mode);
            Controls.Add(_operator);
            Controls.Add(_value1);
            Controls.Add(_value2);
            Controls.Add(_selectValues);
        }

        private void UpdateValueControls()
        {
            var basic = Convert.ToString(_mode.SelectedItem) == "Basic";
            _selectValues.Visible = basic;
            _selectValues.Text = _filter.Values.Count == 0 ? "Select values…" : _filter.Values.Count + " value(s) selected";
            _operator.Visible = !basic;
            _value1.Visible = !basic;
            _value2.Visible = !basic && Convert.ToString(_operator.SelectedItem) == "Between";
        }

        private void ShowValuePicker()
        {
            var selected = new HashSet<string>(_filter.Values, StringComparer.CurrentCultureIgnoreCase);
            using (var dialog = new Form
            {
                Text = "Select " + _filter.Field.Name + " values", Width = 560, Height = 520,
                StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false
            })
            {
                IReadOnlyList<string> available = Array.Empty<string>();
                string lastLoadKey = null;
                var top = new TableLayoutPanel { Dock = DockStyle.Top, Height = 30, ColumnCount = 3, Padding = new Padding(2) };
                top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
                top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 32));
                var search = new TextBox { Dock = DockStyle.Fill };
                search.HandleCreated += (_, __) => SendCueBanner(search, "Search values");
                var order = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                order.Items.AddRange(new object[] { "Ascending", "Descending" });
                order.SelectedIndex = 0;
                var selectedOnly = new CheckBox
                {
                    Appearance = Appearance.Button, Width = 30, Height = 28, Left = 8, Top = 7,
                    FlatStyle = FlatStyle.Flat, Image = CreateSelectedOnlyIcon(), Text = string.Empty,
                    TextAlign = ContentAlignment.MiddleCenter, ImageAlign = ContentAlignment.MiddleCenter,
                    AccessibleName = "Show selected values only"
                };
                selectedOnly.FlatAppearance.BorderSize = 1;
                var clearSelection = new Button
                {
                    Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Image = CreateEraserIcon(),
                    TabStop = false, AccessibleName = "Clear selection"
                };
                clearSelection.FlatAppearance.BorderSize = 0;
                var toolTip = new ToolTip();
                toolTip.SetToolTip(clearSelection, "Clear selection");
                toolTip.SetToolTip(selectedOnly, "Show selected values only");
                top.Controls.Add(search, 0, 0);
                top.Controls.Add(order, 1, 0);
                top.Controls.Add(clearSelection, 2, 0);
                var list = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };
                var bottom = new Panel { Dock = DockStyle.Bottom, Height = 42 };
                var buttons = new FlowLayoutPanel { Dock = DockStyle.Right, Width = 190, FlowDirection = FlowDirection.RightToLeft };
                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
                buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
                bottom.Controls.Add(selectedOnly);
                bottom.Controls.Add(buttons);
                dialog.Controls.Add(list); dialog.Controls.Add(top); dialog.Controls.Add(bottom);
                dialog.AcceptButton = ok; dialog.CancelButton = cancel;

                var populating = false;
                Action populate = () =>
                {
                    populating = true;
                    try
                    {
                        list.Items.Clear();
                        IEnumerable<string> source = selectedOnly.Checked
                            ? selected.OrderBy(v => v, StringComparer.CurrentCultureIgnoreCase)
                            : available;
                        if (selectedOnly.Checked && order.SelectedIndex == 1) source = source.Reverse();
                        foreach (var value in source.Where(v => string.IsNullOrEmpty(search.Text) ||
                                     v.IndexOf(search.Text, StringComparison.CurrentCultureIgnoreCase) >= 0))
                            list.Items.Add(value, selected.Contains(value));
                    }
                    finally { populating = false; }
                };
                Action load = () =>
                {
                    var loadKey = (order.SelectedIndex == 1 ? "DESC" : "ASC") + "\0" + search.Text.Trim();
                    if (string.Equals(lastLoadKey, loadKey, System.StringComparison.Ordinal)) return;
                    lastLoadKey = loadKey;
                    try
                    {
                        Cursor = Cursors.WaitCursor;
                        available = _loadValues(_filter.Field, order.SelectedIndex == 1, search.Text.Trim());
                        populate();
                    }
                    catch (Exception ex)
                    {
                        lastLoadKey = null;
                        MessageBox.Show("Could not load distinct filter values. " + ex.Message,
                            "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally { Cursor = Cursors.Default; }
                };
                list.ItemCheck += (_, e) =>
                {
                    if (populating) return;
                    var value = Convert.ToString(list.Items[e.Index]);
                    if (e.NewValue == CheckState.Checked) selected.Add(value); else selected.Remove(value);
                    if (selectedOnly.Checked) dialog.BeginInvoke(new Action(populate));
                };
                using (var searchTimer = new Timer { Interval = 350 })
                {
                    searchTimer.Tick += (_, __) =>
                    {
                        searchTimer.Stop();
                        if (selectedOnly.Checked) populate(); else load();
                    };
                    search.TextChanged += (_, __) => { searchTimer.Stop(); searchTimer.Start(); };
                order.SelectedIndexChanged += (_, __) => { if (selectedOnly.Checked) populate(); else load(); };
                selectedOnly.CheckedChanged += (_, __) => { searchTimer.Stop(); if (selectedOnly.Checked) populate(); else load(); };
                clearSelection.Click += (_, __) => { selected.Clear(); populate(); };
                load();

                if (dialog.ShowDialog(FindForm()) == DialogResult.OK)
                {
                    _filter.Values = selected.OrderBy(v => v, StringComparer.CurrentCultureIgnoreCase).ToList();
                    _filter.Value = _filter.Values.FirstOrDefault();
                    UpdateValueControls();
                    _changed?.Invoke();
                }
                }
            }
        }

        private static Bitmap CreateEraserIcon()
        {
            var image = new Bitmap(16, 16);
            using (var graphics = Graphics.FromImage(image))
            {
                graphics.Clear(Color.Transparent);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var fill = new SolidBrush(Color.FromArgb(84, 132, 180)))
                using (var end = new SolidBrush(Color.FromArgb(238, 164, 164)))
                using (var outline = new Pen(Color.FromArgb(55, 70, 85), 1.2f))
                {
                    var points = new[] { new Point(3, 11), new Point(9, 3), new Point(14, 8), new Point(8, 14) };
                    graphics.FillPolygon(fill, points);
                    graphics.FillPolygon(end, new[] { new Point(3, 11), new Point(6, 7), new Point(11, 12), new Point(8, 14) });
                    graphics.DrawPolygon(outline, points);
                    graphics.DrawLine(outline, 6, 7, 11, 12);
                }
            }
            return image;
        }

        private static Bitmap CreateSelectedOnlyIcon()
        {
            var image = new Bitmap(16, 16);
            using (var graphics = Graphics.FromImage(image))
            {
                graphics.Clear(Color.Transparent);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var pen = new Pen(Color.FromArgb(55, 90, 125), 1.5f))
                using (var fill = new SolidBrush(Color.FromArgb(84, 132, 180)))
                {
                    graphics.DrawEllipse(pen, 1.5f, 3.5f, 13f, 8f);
                    graphics.FillEllipse(fill, 6, 6, 4, 4);
                    graphics.DrawLines(pen, new[] { new Point(9, 13), new Point(11, 15), new Point(15, 10) });
                }
            }
            return image;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        private static void SendCueBanner(TextBox box, string text) => SendMessage(box.Handle, 0x1501, IntPtr.Zero, text);

        private Control CreateValueControl(string saved, int left, int top)
        {
            if (DaxQueryBuilder.IsDate(_filter.Field))
            {
                DateTime parsed;
                var hasValue = DateTime.TryParse(saved, out parsed);
                var picker = new DateTimePicker
                {
                    Left = left, Top = top, Width = 138, Format = DateTimePickerFormat.Short,
                    ShowCheckBox = true, Checked = hasValue
                };
                if (hasValue) picker.Value = parsed;
                EventHandler update = (_, __) =>
                {
                    if (!_updatingValues)
                        SetValue(picker, picker.Checked ? picker.Value.ToString("yyyy-MM-dd") : null);
                };
                picker.ValueChanged += update;
                return picker;
            }

            var text = new TextBox { Left = left, Top = top, Width = 138, Text = saved ?? string.Empty };
            text.TextChanged += (_, __) => SetValue(text, text.Text);
            return text;
        }

        private void ClearSelection()
        {
            var hadCondition = DaxQueryBuilder.HasCondition(_filter);
            _updatingValues = true;
            try
            {
                _filter.Values.Clear();
                _filter.Value = null;
                _filter.Value2 = null;
                if (_value1 is TextBox firstText) firstText.Text = string.Empty;
                if (_value2 is TextBox secondText) secondText.Text = string.Empty;
                if (_value1 is DateTimePicker firstDate) firstDate.Checked = false;
                if (_value2 is DateTimePicker secondDate) secondDate.Checked = false;
                UpdateValueControls();
            }
            finally { _updatingValues = false; }
            if (hadCondition) _changed?.Invoke(); else _stateChanged?.Invoke();
        }

        private void SetValue(Control sender, string value)
        {
            if (_updatingValues) return;
            if (ReferenceEquals(sender, _value2)) _filter.Value2 = value;
            else if (_value1 == null || ReferenceEquals(sender, _value1)) _filter.Value = value;
            else _filter.Value2 = value;
            _changed?.Invoke();
        }

        private static string[] Operators(SemanticField field)
        {
            if (DaxQueryBuilder.IsDate(field)) return new[]
                { "Equals", "Not Equals", "Greater Than", "Greater Than Or Equal", "Less Than", "Less Than Or Equal", "Between" };
            if (DaxQueryBuilder.IsNumeric(field))
                return new[] { "Equals", "Not Equals", "Greater Than", "Greater Than Or Equal", "Less Than", "Less Than Or Equal", "Between" };
            return new[] { "Equals", "Not Equals", "Contains", "Starts With", "Ends With" };
        }
    }
}
