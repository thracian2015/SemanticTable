using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace SemanticTable
{
    [ComVisible(true)]
    public sealed class FieldsPane : UserControl
    {
        private const int EmSetCueBanner = 0x1501;
        private const int TvmSetItemW = 0x113F;
        private const uint TvifState = 0x0008;
        private const uint TvisStateImageMask = 0xF000;
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
        private static extern IntPtr SendTreeMessage(IntPtr hWnd, int msg, IntPtr wParam, ref TvItem item);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct TvItem
        {
            public uint mask;
            public IntPtr hItem;
            public uint state;
            public uint stateMask;
            public IntPtr pszText;
            public int cchTextMax;
            public int iImage;
            public int iSelectedImage;
            public int cChildren;
            public IntPtr lParam;
        }

        private readonly TextBox _search = new TextBox { Dock = DockStyle.Top };
        private readonly Button _workspace = new Button
        {
            Text = "", Dock = DockStyle.Fill, Margin = new Padding(2, 0, 0, 1),
            FlatStyle = FlatStyle.Flat,
            AccessibleName = "Connection settings"
        };
        private readonly Button _refreshFields = new Button
        {
            Text = "", Dock = DockStyle.Fill, Margin = new Padding(2, 0, 0, 1),
            FlatStyle = FlatStyle.Flat, AccessibleName = "Refresh field list"
        };
        private readonly TreeView _tree = new TreeView { Dock = DockStyle.Fill, CheckBoxes = true, HideSelection = false };
        private readonly Button _apply = new Button { Text = "Apply", AutoSize = true };
        private readonly CheckBox _deferUpdate = new CheckBox { Text = "Defer update", AutoSize = true, Checked = true, Margin = new Padding(3, 7, 8, 3) };
        private readonly Timer _autoApplyTimer = new Timer { Interval = 450 };
        private readonly Label _status = new Label
        {
            AutoSize = false,
            Height = 24,
            AutoEllipsis = true,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.DimGray,
            Text = "Select a connected table."
        };
        private readonly FlowLayoutPanel _filters = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, AllowDrop = true };
        private readonly HashSet<string> _checkedKeys = new HashSet<string>();
        private readonly ImageList _icons = new ImageList { ImageSize = new Size(16, 16), ColorDepth = ColorDepth.Depth32Bit };
        private readonly ContextMenuStrip _fieldMenu = new ContextMenuStrip();
        private readonly ExcelAdoMetadataProvider _metadata = new ExcelAdoMetadataProvider();
        private SemanticField _contextField;
        private ConnectedTableContext _context;
        private TableDefinition _definition;
        private IReadOnlyList<SemanticField> _allFields = Array.Empty<SemanticField>();
        private IReadOnlyList<SemanticHierarchy> _hierarchies = Array.Empty<SemanticHierarchy>();
        private bool _syncingChecks;
        private bool _loadingState;
        private bool _applying;
        private bool _contextMismatch;
        private Excel.Application _excelApp;

        public FieldsPane()
        {
            _search.Text = "";
            _search.HandleCreated += (_, __) => SendMessage(_search.Handle, EmSetCueBanner, IntPtr.Zero, "Search fields");
            _icons.Images.Add("table", CreateTableIcon());
            _icons.Images.Add("folder", CreateFolderIcon());
            _icons.Images.Add("hierarchy", CreateHierarchyIcon());
            _icons.Images.Add("column", CreateColumnIcon());
            _icons.Images.Add("measure", CreateMeasureIcon());
            _workspace.Image = CreateConnectionIcon();
            _workspace.FlatAppearance.BorderSize = 0;
            _refreshFields.Image = CreateRefreshIcon();
            _refreshFields.FlatAppearance.BorderSize = 0;
            var toolTip = new ToolTip();
            toolTip.SetToolTip(_workspace, "Connection settings");
            toolTip.SetToolTip(_refreshFields, "Refresh fields for the active connected table");
            _tree.ImageList = _icons;
            _tree.TreeViewNodeSorter = new MetadataNodeComparer();
            _fieldMenu.Items.Add("Add to filters", null, (_, __) => AddFilter(_contextField));
            try
            {
                _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                _excelApp.SheetSelectionChange += OnExcelSheetSelectionChange;
            }
            catch { _excelApp = null; }
            Disposed += (_, __) =>
            {
                try { if (_excelApp != null) _excelApp.SheetSelectionChange -= OnExcelSheetSelectionChange; } catch { }
                _fieldMenu.Dispose();
                _autoApplyTimer.Dispose();
            };
            var bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                Padding = new Padding(6),
                Margin = Padding.Empty,
                ColumnCount = 1,
                RowCount = 2
            };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            bottom.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            var bottomActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                WrapContents = false
            };
            bottomActions.Controls.Add(_deferUpdate);
            bottomActions.Controls.Add(_apply);
            _status.Dock = DockStyle.Fill;
            _status.Margin = new Padding(3, 0, 3, 0);
            bottom.Controls.Add(bottomActions, 0, 0);
            bottom.Controls.Add(_status, 0, 1);
            var filterHost = new Panel { Dock = DockStyle.Fill };
            filterHost.Controls.Add(_filters);
            filterHost.Controls.Add(new Label
            {
                Text = "Filters — drag a column here (filter fields need not appear in output)",
                Dock = DockStyle.Top, Height = 34, Padding = new Padding(6, 8, 3, 3), ForeColor = Color.DimGray
            });

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 6,
                Panel1MinSize = 80,
                Panel2MinSize = 80
            };
            split.Panel1.Controls.Add(_tree);
            split.Panel2.Controls.Add(filterHost);

            var searchRow = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                ColumnCount = 3,
                RowCount = 1
            };
            searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 62));
            searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 8));
            searchRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var searchButtons = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, Height = 27, Margin = Padding.Empty,
                Padding = Padding.Empty, ColumnCount = 2, RowCount = 1
            };
            searchButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 31));
            searchButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 31));
            searchButtons.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _refreshFields.Dock = DockStyle.Fill;
            _refreshFields.Margin = Padding.Empty;
            _workspace.Dock = DockStyle.Fill;
            _workspace.Margin = Padding.Empty;
            searchButtons.Controls.Add(_refreshFields, 0, 0);
            searchButtons.Controls.Add(_workspace, 1, 0);
            _search.Dock = DockStyle.Fill;
            searchRow.Controls.Add(_search, 0, 0);
            searchRow.Controls.Add(searchButtons, 1, 0);

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 27));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(searchRow, 0, 0);
            layout.Controls.Add(split, 0, 1);
            layout.Controls.Add(bottom, 0, 2);
            Controls.Add(layout);
            _search.TextChanged += (_, __) => { if (!_contextMismatch) PopulateTree(); };
            _apply.Click += (_, __) => Apply();
            _autoApplyTimer.Tick += (_, __) => { _autoApplyTimer.Stop(); if (!_deferUpdate.Checked) Apply(); };
            _deferUpdate.CheckedChanged += (_, __) =>
            {
                if (_loadingState || _definition == null) return;
                try
                {
                    if (_deferUpdate.Checked) _autoApplyTimer.Stop();
                    _definition.DeferUpdate = _deferUpdate.Checked;
                    _apply.Enabled = _deferUpdate.Checked;
                    SaveDefinition();
                    DiagnosticLog.Write("Defer update changed to " + _deferUpdate.Checked + ".");
                    if (!_deferUpdate.Checked) ScheduleAutoApply();
                }
                catch (Exception ex)
                {
                    DiagnosticLog.Write("Defer update handler failed: " + ex);
                    ShowError(new InvalidOperationException(
                        "Could not save the Defer update setting. " + ExceptionDetails(ex) +
                        "\r\n\r\nDiagnostic log: " + DiagnosticLog.PathName, ex));
                }
            };
            _workspace.Click += (_, __) => ChangeDisplayedConnection();
            _refreshFields.Click += (_, __) => RefreshActiveTableFields();
            _tree.BeforeCheck += (_, e) =>
            {
                if (!(e.Node.Tag is SemanticField)) e.Cancel = true;
            };
            _tree.AfterCheck += (_, e) =>
            {
                var field = e.Node.Tag as SemanticField;
                if (field == null) return;
                if (e.Node.Checked) _checkedKeys.Add(field.Key); else _checkedKeys.Remove(field.Key);
                if (_syncingChecks) return;
                try
                {
                    _syncingChecks = true;
                    SetMatchingNodeChecks(_tree.Nodes, field.Key, e.Node.Checked, e.Node);
                }
                finally { _syncingChecks = false; }
                ScheduleAutoApply();
            };
            _tree.ItemDrag += (_, e) =>
            {
                var field = (e.Item as TreeNode)?.Tag as SemanticField;
                if (field != null) DoDragDrop(field, DragDropEffects.Copy);
            };
            _tree.NodeMouseClick += (_, e) =>
            {
                if (e.Button != MouseButtons.Right) return;
                var field = e.Node.Tag as SemanticField;
                if (field == null || field.Kind != SemanticFieldKind.Column) return;
                _tree.SelectedNode = e.Node;
                _contextField = field;
                _fieldMenu.Show(_tree, e.Location);
            };
            _filters.DragEnter += (_, e) =>
            {
                e.Effect = e.Data.GetDataPresent(typeof(SemanticField)) ? DragDropEffects.Copy : DragDropEffects.None;
            };
            _filters.DragDrop += (_, e) => AddFilter(e.Data.GetData(typeof(SemanticField)) as SemanticField);
            _filters.ClientSizeChanged += (_, __) => ResizeFilterRows();
        }

        public void AttachToSelection()
        {
            var step = "starting";
            var previousContext = _context;
            var previousDefinition = _definition;
            var previousFields = _allFields;
            var previousHierarchies = _hierarchies;
            var previousCheckedKeys = new HashSet<string>(_checkedKeys);
            var previousStatus = _status.Text;
            try
            {
                _autoApplyTimer.Stop();
                Cursor = Cursors.WaitCursor;
                step = "getting the Excel application";
                var app = (Excel.Application)ExcelDnaUtil.Application;
                step = "locating or creating the connected table";
                var workbook = app.ActiveWorkbook;
                _context = ExcelConnectionService.GetActiveSheetConnectedTable(app);
                if (_context == null)
                {
                    var entered = PromptForConnectionString(NewConnectionTemplate(), true);
                    if (entered == null) throw new InvalidOperationException("Connection entry was canceled.");
                    var connection = ValidateConnectionString(entered);
                    _context = ExcelConnectionService.CreateConnectedTable(app, connection);
                    StateStore.SaveConnectionString(workbook, _context.Table.Name, connection);
                }
                step = "resolving the semantic-model connection";
                var savedConnection = StateStore.LoadConnectionString(workbook, _context.Table.Name);
                if (string.IsNullOrWhiteSpace(savedConnection))
                {
                    var entered = PromptForConnectionString(ConnectionTemplate(_context.ConnectionString), false);
                    if (entered == null) throw new InvalidOperationException("Connection entry was canceled.");
                    savedConnection = ValidateConnectionString(entered);
                    StateStore.SaveConnectionString(workbook, _context.Table.Name, savedConnection);
                }
                _context.ConnectionString = savedConnection;
                var datasetId = ExcelConnectionService.GetModelIdentity(savedConnection);
                step = "loading saved field state";
                _definition = StateStore.Load(app.ActiveWorkbook, _context.Table.Name);
                var resetModelState = false;
                if (_definition.Version < 2)
                {
                    if (_definition.RowLimit == 10000) _definition.RowLimit = 500000;
                    _definition.Version = 2;
                }
                if (_definition.Version < 3)
                {
                    if (_definition.Filters != null)
                        foreach (var filter in _definition.Filters)
                            filter.Mode = string.Equals(filter.Operator, "Equals", StringComparison.OrdinalIgnoreCase)
                                ? "Basic" : "Advanced";
                    _definition.Version = 3;
                }
                if (_definition.Version < 4 ||
                    !string.Equals(_definition.DatasetId, datasetId, StringComparison.OrdinalIgnoreCase))
                {
                    _definition.Fields = new List<SemanticField>();
                    _definition.Filters = new List<FieldFilter>();
                    _definition.DatasetId = datasetId;
                    _definition.Version = 4;
                    resetModelState = true;
                }
                if (_definition.Version < 5)
                {
                    _definition.DeferUpdate = true;
                    _definition.Version = 5;
                    resetModelState = true;
                }
                _loadingState = true;
                _deferUpdate.Checked = _definition.DeferUpdate;
                _apply.Enabled = _definition.DeferUpdate;
                _loadingState = false;
                _status.Text = "Loading model fields…";
                step = "reading semantic-model metadata";
                _allFields = _metadata.Load(_context);
                _hierarchies = _metadata.Hierarchies;
                step = "matching fields to the connected table query";
                var queryFields = DaxQueryBuilder.FindReferencedFields(_context.CommandText, _allFields);
                if (queryFields.Count > 0)
                    _definition.Fields = queryFields;
                if (_definition.Filters == null) _definition.Filters = new List<FieldFilter>();
                var removedFilters = ReconcileFilters();
                if (resetModelState || removedFilters > 0) StateStore.Save(app.ActiveWorkbook, _definition);
                _checkedKeys.Clear();
                foreach (var field in _definition.Fields) _checkedKeys.Add(field.Key);
                step = "populating the Fields pane";
                PopulateTree();
                RenderFilters();
                _contextMismatch = false;
                SetFieldListEnabled(true);
                _status.Text = $"{_definition.Fields.Count} selected / {_allFields.Count} fields";
            }
            catch (Exception ex)
            {
                // Switching tables is transactional. If the candidate table/model
                // cannot be opened, keep the pane bound to the table it still displays.
                _context = previousContext;
                _definition = previousDefinition;
                _allFields = previousFields;
                _hierarchies = previousHierarchies;
                _checkedKeys.Clear();
                foreach (var key in previousCheckedKeys) _checkedKeys.Add(key);
                _loadingState = true;
                _deferUpdate.Checked = previousDefinition?.DeferUpdate ?? true;
                _apply.Enabled = previousDefinition != null && previousDefinition.DeferUpdate;
                _loadingState = false;
                if (previousDefinition != null)
                {
                    PopulateTree();
                    RenderFilters();
                }
                _status.Text = previousStatus;
                if (_contextMismatch) SetFieldListEnabled(false);
                ShowError(new InvalidOperationException($"Semantic Table failed while {step}.\r\n\r\n{ex.Message}", ex));
                _status.Text = previousStatus;
            }
            finally { Cursor = Cursors.Default; }
        }

        private static string PromptForConnectionString(string current, bool creatingTable)
        {
            using (var dialog = new Form
            {
                Text = creatingTable ? "Create Semantic Table connection" : "Semantic Table connection",
                Width = 820,
                Height = 285,
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            })
            {
                var label = new Label
                {
                    Left = 12, Top = 12, Width = 780, Height = 52,
                    Text = creatingTable
                        ? "Enter the complete MSOLAP connection string for the semantic model. Replace <workspace> and <semantic-model> with actual values."
                        : "Edit the complete connection string used by Semantic Table. Replace the <workspace> placeholder with the Power BI workspace name. This setting is saved for this table."
                };
                var textBox = new TextBox
                {
                    Left = 12, Top = 66, Width = 780, Height = 125, Text = current ?? "",
                    Multiline = true, ScrollBars = ScrollBars.Vertical, AcceptsReturn = true
                };
                var ok = new Button { Text = "OK", Left = 636, Top = 205, Width = 75, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "Cancel", Left = 717, Top = 205, Width = 75, DialogResult = DialogResult.Cancel };
                dialog.Controls.Add(label); dialog.Controls.Add(textBox); dialog.Controls.Add(ok); dialog.Controls.Add(cancel);
                dialog.AcceptButton = ok; dialog.CancelButton = cancel;
                return dialog.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : null;
            }
        }

        private static string ConnectionTemplate(string nativeConnection)
        {
            var value = nativeConnection ?? string.Empty;
            if (value.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase)) value = value.Substring(6);
            if (ExcelConnectionService.UsesExcelPrivatePowerBiEndpoint(value))
                value = ExcelConnectionService.ReplaceDataSource(value,
                    "powerbi://api.powerbi.com/v1.0/myorg/<workspace>");
            return value;
        }

        private static string NewConnectionTemplate() =>
            "Provider=MSOLAP.8;Data Source=powerbi://api.powerbi.com/v1.0/myorg/<workspace>;" +
            "Initial Catalog=<semantic-model>";

        private static string ValidateConnectionString(string connection)
        {
            var value = (connection ?? string.Empty).Trim();
            if (value.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase)) value = value.Substring(6);
            if (string.IsNullOrWhiteSpace(value)) throw new InvalidOperationException("Enter a connection string.");
            if (value.IndexOf("<workspace>", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("<semantic-model>", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException("Replace all connection placeholders with actual values.");
            if (string.IsNullOrWhiteSpace(ExcelConnectionService.GetProperty(value, "Data Source")))
                throw new InvalidOperationException("The connection string must contain Data Source.");
            if (string.IsNullOrWhiteSpace(ExcelConnectionService.GetProperty(value, "Initial Catalog")))
                throw new InvalidOperationException("The connection string must contain Initial Catalog.");
            if (string.IsNullOrWhiteSpace(ExcelConnectionService.GetProperty(value, "Provider")))
                value = "Provider=MSOLAP.8;" + value;
            return value;
        }

        internal void ShowSettingsDialog()
        {
            if (_context == null || _definition == null)
                throw new InvalidOperationException("Open Semantic Table Fields for a connected table first.");

            using (var dialog = new Form
            {
                Text = "Semantic Table Settings", Width = 480, Height = 250,
                StartPosition = FormStartPosition.CenterScreen, MinimizeBox = false, MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            })
            {
                var grid = new DataGridView
                {
                    Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                    AllowUserToResizeRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                };
                grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "Name", ReadOnly = true });
                grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Value", HeaderText = "Value" });
                grid.Rows.Add("Row Limit", _definition.RowLimit.ToString());

                var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft };
                var ok = new Button { Text = "OK", Width = 80 };
                var cancel = new Button { Text = "Cancel", Width = 80, DialogResult = DialogResult.Cancel };
                buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
                dialog.Controls.Add(grid); dialog.Controls.Add(buttons);
                dialog.CancelButton = cancel;
                ok.Click += (_, __) =>
                {
                    int value;
                    if (!int.TryParse(Convert.ToString(grid.Rows[0].Cells["Value"].Value), out value) || value <= 0)
                    {
                        MessageBox.Show("Row Limit must be a positive integer.", "Semantic Table Settings",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    _definition.RowLimit = value;
                    SaveDefinition();
                    dialog.DialogResult = DialogResult.OK;
                    dialog.Close();
                };
                if (dialog.ShowDialog(FindForm()) == DialogResult.OK) ScheduleAutoApply();
            }
        }

        private void ChangeDisplayedConnection()
        {
            if (_contextMismatch)
            {
                MessageBox.Show("Click Refresh to bind the Fields pane to the active connected table before editing its connection.",
                    "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_context == null)
            {
                MessageBox.Show("Open Semantic Table Fields for a connected table first.", "Semantic Table",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                Cursor = Cursors.WaitCursor;
                var workbook = (Excel.Workbook)((Excel.Worksheet)_context.Table.Parent).Parent;
                var entered = PromptForConnectionString(_context.ConnectionString, false);
                if (entered == null) return;
                var connection = ValidateConnectionString(entered);
                StateStore.SaveConnectionString(workbook, _context.Table.Name, connection);
                var previousDatasetId = _definition.DatasetId;
                _context.ConnectionString = connection;
                var datasetId = ExcelConnectionService.GetModelIdentity(connection);
                _status.Text = "Loading model fields…";
                _allFields = _metadata.Load(_context);
                _hierarchies = _metadata.Hierarchies;
                if (!string.Equals(previousDatasetId, datasetId, StringComparison.OrdinalIgnoreCase))
                {
                    _definition.Fields = new List<SemanticField>();
                    _definition.Filters = new List<FieldFilter>();
                    _definition.DatasetId = datasetId;
                }
                ReconcileFilters();
                if (string.Equals(previousDatasetId, datasetId, StringComparison.OrdinalIgnoreCase))
                {
                    var queryFields = DaxQueryBuilder.FindReferencedFields(_context.CommandText, _allFields);
                    if (queryFields.Count > 0) _definition.Fields = queryFields;
                }
                StateStore.Save(workbook, _definition);
                _checkedKeys.Clear();
                foreach (var field in _definition.Fields) _checkedKeys.Add(field.Key);
                PopulateTree();
                RenderFilters();
                _status.Text = $"{_definition.Fields.Count} selected / {_allFields.Count} fields";
            }
            catch (Exception ex) { ShowError(ex); }
            finally { Cursor = Cursors.Default; }
        }

        private void RefreshActiveTableFields()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (ExcelConnectionService.GetActiveSheetConnectedTable(app) == null)
                {
                    MessageBox.Show("The active worksheet does not contain a semantic-model connected table.",
                        "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                AttachToSelection();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void OnExcelSheetSelectionChange(object sheet, Excel.Range target)
        {
            if (_context == null || _excelApp == null) return;
            try
            {
                var active = ExcelConnectionService.GetConnectedTableAtActiveCell(_excelApp);
                if (active == null) return;
                if (IsSameTable(_context, active))
                {
                    if (_contextMismatch)
                    {
                        _contextMismatch = false;
                        SetFieldListEnabled(true);
                        _status.Text = $"{_definition.Fields.Count} selected / {_allFields.Count} fields";
                    }
                    return;
                }
                _contextMismatch = true;
                _autoApplyTimer.Stop();
                SetFieldListEnabled(false);
                _status.Text = "Metadata mismatch-click Refresh";
            }
            catch { }
        }

        private void SetFieldListEnabled(bool enabled)
        {
            // Keep the complete toolbar row enabled. Disabling its fill control in
            // the Excel task-pane host can paint over the adjacent icon buttons.
            _search.Enabled = true;
            _tree.Enabled = enabled;
            _filters.Enabled = enabled;
            _deferUpdate.Enabled = enabled;
            _workspace.Enabled = true;
            _apply.Enabled = enabled && !_applying && _deferUpdate.Checked;
            _refreshFields.Enabled = true;
            _workspace.Visible = true;
            _refreshFields.Visible = true;
            _workspace.Parent?.PerformLayout();
        }

        private static bool IsSameTable(ConnectedTableContext left, ConnectedTableContext right)
        {
            try
            {
                var leftSheet = (Excel.Worksheet)left.Table.Parent;
                var rightSheet = (Excel.Worksheet)right.Table.Parent;
                var leftBook = (Excel.Workbook)leftSheet.Parent;
                var rightBook = (Excel.Workbook)rightSheet.Parent;
                return string.Equals(left.Table.Name, right.Table.Name, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(leftSheet.Name, rightSheet.Name, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(leftBook.Name, rightBook.Name, StringComparison.OrdinalIgnoreCase);
            }
            catch { return ReferenceEquals(left.Table, right.Table); }
        }

        private int ReconcileFilters()
        {
            if (_definition.Filters == null)
            {
                _definition.Filters = new List<FieldFilter>();
                return 0;
            }
            var removed = 0;
            var valid = new List<FieldFilter>();
            foreach (var filter in _definition.Filters)
            {
                var current = filter?.Field == null ? null : _allFields.FirstOrDefault(f => f.Key == filter.Field.Key);
                if (current == null)
                {
                    removed++;
                    continue;
                }
                filter.Field = current;
                valid.Add(filter);
            }
            _definition.Filters = valid;
            return removed;
        }

        private void PopulateTree()
        {
            var wasLoading = _loadingState;
            _loadingState = true;
            var filter = _search.Text.Trim();
            _tree.BeginUpdate();
            _tree.Nodes.Clear();
            foreach (var group in _allFields.Where(f => filter.Length == 0 ||
                f.Name.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) >= 0 ||
                f.Table.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) >= 0 ||
                (f.DisplayFolder ?? "").IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) >= 0).GroupBy(f => f.Table))
            {
                var tableNode = new TreeNode(group.Key) { ImageKey = "table", SelectedImageKey = "table" };
                foreach (var field in group)
                {
                    var parent = DisplayFolderNode(tableNode, field.DisplayFolder);
                    parent.Nodes.Add(new TreeNode(field.Display)
                    {
                        Tag = field,
                        Checked = _checkedKeys.Contains(field.Key),
                        ImageKey = field.Kind == SemanticFieldKind.Measure ? "measure" : "column",
                        SelectedImageKey = field.Kind == SemanticFieldKind.Measure ? "measure" : "column"
                    });
                }
                foreach (var hierarchy in _hierarchies.Where(h => string.Equals(h.Table, group.Key, StringComparison.OrdinalIgnoreCase)))
                {
                    if (filter.Length > 0 && hierarchy.Name.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) < 0 &&
                        (hierarchy.DisplayFolder ?? "").IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) < 0 &&
                        !hierarchy.Levels.Any(l => l.Name.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) >= 0)) continue;
                    var parent = DisplayFolderNode(tableNode, hierarchy.DisplayFolder);
                    var hierarchyNode = new TreeNode(hierarchy.Name)
                    {
                        Tag = hierarchy,
                        ImageKey = "hierarchy",
                        SelectedImageKey = "hierarchy"
                    };
                    foreach (var level in hierarchy.Levels)
                        hierarchyNode.Nodes.Add(new TreeNode(level.Display)
                        {
                            Tag = level,
                            Checked = _checkedKeys.Contains(level.Key),
                            ImageKey = "column",
                            SelectedImageKey = "column"
                        });
                    parent.Nodes.Add(hierarchyNode);
                }
                _tree.Nodes.Add(tableNode);
            }
            _tree.Sort();
            HideNonFieldCheckboxes(_tree.Nodes);
            _tree.EndUpdate();
            _loadingState = wasLoading;
        }

        private void HideNonFieldCheckboxes(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (!(node.Tag is SemanticField))
                {
                    var item = new TvItem
                    {
                        mask = TvifState,
                        hItem = node.Handle,
                        stateMask = TvisStateImageMask,
                        state = 0
                    };
                    SendTreeMessage(_tree.Handle, TvmSetItemW, IntPtr.Zero, ref item);
                }
                if (node.Nodes.Count > 0) HideNonFieldCheckboxes(node.Nodes);
            }
        }

        private static void SetMatchingNodeChecks(TreeNodeCollection nodes, string key, bool isChecked, TreeNode except)
        {
            foreach (TreeNode node in nodes)
            {
                var field = node.Tag as SemanticField;
                if (!ReferenceEquals(node, except) && field != null && field.Key == key) node.Checked = isChecked;
                if (node.Nodes.Count > 0) SetMatchingNodeChecks(node.Nodes, key, isChecked, except);
            }
        }

        private static TreeNode DisplayFolderNode(TreeNode tableNode, string displayFolder)
        {
            var parent = tableNode;
            foreach (var part in (displayFolder ?? "").Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var name = part.Trim();
                if (name.Length == 0) continue;
                var existing = parent.Nodes.Cast<TreeNode>().FirstOrDefault(n =>
                    n.Tag == null && string.Equals(n.Text, name, StringComparison.CurrentCultureIgnoreCase));
                if (existing == null)
                {
                    existing = new TreeNode(name) { ImageKey = "folder", SelectedImageKey = "folder" };
                    parent.Nodes.Add(existing);
                }
                parent = existing;
            }
            return parent;
        }

        private void Apply()
        {
            if (_context == null || _applying || !FiltersAreComplete()) return;
            if (_definition == null || !string.Equals(_definition.ExcelTableName, _context.Table.Name, StringComparison.OrdinalIgnoreCase))
            {
                ShowError(new InvalidOperationException(
                    "The Fields pane and connected-table context do not match. Reopen Fields for the intended table before applying changes."));
                return;
            }
            string dax = null;
            try
            {
                _applying = true;
                Cursor = Cursors.WaitCursor;
                _apply.Enabled = false;
                var selected = _allFields.Where(f => _checkedKeys.Contains(f.Key)).ToList();
                dax = DaxQueryBuilder.Build(selected, _definition.Filters, _definition.RowLimit);
                _status.Text = "Refreshing…";
                ExcelConnectionService.ApplyAndRefresh(_context, dax, selected.Count);
                _definition.Fields = selected;
                DiagnosticLog.Write("Saving table definition after successful refresh.");
                SaveDefinition();
                DiagnosticLog.Write("Table definition saved successfully.");
                _status.Text = $"{selected.Count} selected, {_definition.Filters.Count} filters — updated {DateTime.Now:t}";
            }
            catch (Exception ex)
            {
                var message = string.IsNullOrWhiteSpace(dax)
                    ? ExceptionDetails(ex)
                    : "Excel failed while updating the connected table.\r\n\r\n" + ExceptionDetails(ex);
                if (!string.IsNullOrWhiteSpace(dax) && message.IndexOf("Generated DAX:", StringComparison.OrdinalIgnoreCase) < 0)
                    message += "\r\n\r\nGenerated DAX:\r\n" + dax;
                ShowError(new InvalidOperationException(message, ex));
            }
            finally
            {
                _applying = false;
                _apply.Enabled = _deferUpdate.Checked;
                Cursor = Cursors.Default;
            }
        }

        private bool FiltersAreComplete()
        {
            if (_definition?.Filters == null) return true;
            return _definition.Filters.All(filter =>
            {
                if (filter?.Field == null) return false;
                if (string.Equals(filter.Mode, "Basic", StringComparison.OrdinalIgnoreCase))
                    return true;
                if (string.IsNullOrWhiteSpace(filter.Value))
                    return string.IsNullOrWhiteSpace(filter.Value2);
                return !string.Equals(filter.Operator, "Between", StringComparison.OrdinalIgnoreCase) ||
                       !string.IsNullOrWhiteSpace(filter.Value2);
            });
        }

        private void ScheduleAutoApply()
        {
            if (_loadingState || _definition == null || _deferUpdate.Checked || _checkedKeys.Count == 0 || !FiltersAreComplete()) return;
            _autoApplyTimer.Stop();
            _autoApplyTimer.Start();
        }

        private void SaveDefinition()
        {
            if (_definition == null || _context == null) return;
            var workbook = (Excel.Workbook)((Excel.Worksheet)_context.Table.Parent).Parent;
            StateStore.Save(workbook, _definition);
        }

        private static string ExceptionDetails(Exception error)
        {
            var messages = new List<string>();
            for (var current = error; current != null; current = current.InnerException)
                if (!string.IsNullOrWhiteSpace(current.Message) && !messages.Contains(current.Message))
                    messages.Add(current.Message);
            return string.Join("\r\n", messages);
        }

        private void AddFilter(SemanticField field)
        {
            if (field == null) return;
            if (field.Kind == SemanticFieldKind.Measure)
            {
                MessageBox.Show("Measure filters are not supported in this version. Drag a model column into Filters.",
                    "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_definition.Filters.Any(f => f.Field != null && f.Field.Key == field.Key)) return;
            _definition.Filters.Add(new FieldFilter { Field = field, Mode = "Basic", Operator = "Equals" });
            RenderFilters();
            SaveFilterUiState();
        }

        private void RemoveFilter(FieldFilter filter)
        {
            _definition.Filters.Remove(filter);
            RenderFilters();
            ScheduleAutoApply();
        }

        private void RenderFilters()
        {
            _filters.SuspendLayout();
            _filters.Controls.Clear();
            foreach (var filter in _definition.Filters)
            {
                if (filter.Field == null) continue;
                var current = _allFields.FirstOrDefault(f => f.Key == filter.Field.Key);
                if (current != null) filter.Field = current;
                var row = new FilterRow(filter, RemoveFilter, LoadFilterValues, ScheduleAutoApply, SaveFilterUiState)
                {
                    Width = FilterRowWidth()
                };
                _filters.Controls.Add(row);
            }
            if (_definition.Filters.Count == 0)
                _filters.Controls.Add(new Label
                {
                    Text = "Drag a column from the Fields tree into this area.",
                    AutoSize = true, ForeColor = Color.Gray, Padding = new Padding(8, 12, 3, 3)
                });
            _filters.ResumeLayout();
            ResizeFilterRows();
            _filters.PerformLayout();
        }

        private int FilterRowWidth()
        {
            var scrollbar = _filters.VerticalScroll.Visible ? SystemInformation.VerticalScrollBarWidth : 0;
            return Math.Max(100, _filters.ClientSize.Width - scrollbar);
        }

        private void ResizeFilterRows()
        {
            var width = FilterRowWidth();
            foreach (Control control in _filters.Controls)
                if (control is FilterRow row) row.Width = width;
        }

        private IReadOnlyList<string> LoadFilterValues(SemanticField field, bool descending, string search) =>
            _metadata.LoadDistinctValues(_context, field, 500, descending, search);

        private void SaveFilterUiState()
        {
            try { SaveDefinition(); }
            catch (Exception ex)
            {
                DiagnosticLog.Write("Could not save filter UI state: " + ex);
                ShowError(new InvalidOperationException("Could not save the filter mode. " + ExceptionDetails(ex), ex));
            }
        }

        private static Bitmap CreateConnectionIcon()
        {
            var image = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.FromArgb(55, 90, 125), 1.6f))
            {
                g.Clear(Color.Transparent);
                g.DrawLine(pen, 5, 1, 5, 5);
                g.DrawLine(pen, 11, 1, 11, 5);
                g.DrawRectangle(pen, 3, 5, 10, 5);
                g.DrawArc(pen, 5, 8, 6, 6, 0, 180);
                g.DrawLine(pen, 8, 13, 8, 16);
            }
            return image;
        }

        private static Bitmap CreateRefreshIcon()
        {
            var image = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.FromArgb(16, 124, 65), 1.7f))
            using (var brush = new SolidBrush(Color.FromArgb(16, 124, 65)))
            {
                g.Clear(Color.Transparent);
                g.DrawArc(pen, 2, 2, 11, 11, 35, 285);
                g.FillPolygon(brush, new[] { new Point(12, 1), new Point(15, 5), new Point(10, 5) });
            }
            return image;
        }

        private static Bitmap CreateTableIcon()
        {
            var image = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.DimGray))
            {
                g.Clear(Color.Transparent);
                g.DrawRectangle(pen, 1, 2, 13, 12);
                g.DrawLine(pen, 1, 6, 14, 6);
                g.DrawLine(pen, 5, 2, 5, 14);
                g.DrawLine(pen, 10, 2, 10, 14);
            }
            return image;
        }

        private static Bitmap CreateColumnIcon()
        {
            var image = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.SteelBlue))
            {
                g.Clear(Color.Transparent);
                g.DrawRectangle(pen, 3, 1, 9, 14);
                g.DrawLine(pen, 3, 5, 12, 5);
                g.DrawLine(pen, 3, 9, 12, 9);
            }
            return image;
        }

        private static Bitmap CreateFolderIcon()
        {
            var image = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.Goldenrod))
            using (var brush = new SolidBrush(Color.FromArgb(255, 224, 130)))
            {
                g.Clear(Color.Transparent);
                g.FillRectangle(brush, 1, 5, 14, 9);
                g.FillRectangle(brush, 2, 3, 6, 3);
                g.DrawRectangle(pen, 1, 5, 14, 9);
                g.DrawLine(pen, 2, 3, 8, 3);
                g.DrawLine(pen, 8, 3, 10, 5);
            }
            return image;
        }

        private static Bitmap CreateMeasureIcon()
        {
            var image = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.SeaGreen))
            using (var brush = new SolidBrush(Color.SeaGreen))
            {
                g.Clear(Color.Transparent);
                g.DrawRectangle(pen, 2, 1, 11, 14);
                g.DrawRectangle(pen, 4, 3, 7, 3);
                g.FillRectangle(brush, 4, 8, 2, 2);
                g.FillRectangle(brush, 8, 8, 2, 2);
                g.FillRectangle(brush, 4, 12, 2, 2);
                g.FillRectangle(brush, 8, 12, 2, 2);
            }
            return image;
        }

        private static Bitmap CreateHierarchyIcon()
        {
            var image = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.MediumPurple))
            using (var brush = new SolidBrush(Color.MediumPurple))
            {
                g.Clear(Color.Transparent);
                g.DrawLine(pen, 4, 3, 4, 12);
                g.DrawLine(pen, 4, 6, 11, 6);
                g.DrawLine(pen, 4, 11, 11, 11);
                g.FillEllipse(brush, 1, 1, 6, 6);
                g.FillEllipse(brush, 9, 4, 6, 6);
                g.FillEllipse(brush, 9, 9, 6, 6);
            }
            return image;
        }

        private void ShowError(Exception ex)
        {
            _status.Text = "Unable to update table";
            MessageBox.Show(ex.Message, "Semantic Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private sealed class MetadataNodeComparer : IComparer
        {
            public int Compare(object x, object y)
            {
                var leftNode = x as TreeNode;
                var rightNode = y as TreeNode;
                var hierarchy = leftNode?.Parent?.Tag as SemanticHierarchy;
                if (hierarchy != null && ReferenceEquals(leftNode.Parent, rightNode?.Parent))
                {
                    var leftField = leftNode.Tag as SemanticField;
                    var rightField = rightNode.Tag as SemanticField;
                    var leftOrdinal = hierarchy.Levels.IndexOf(leftField);
                    var rightOrdinal = hierarchy.Levels.IndexOf(rightField);
                    if (leftOrdinal >= 0 && rightOrdinal >= 0) return leftOrdinal.CompareTo(rightOrdinal);
                }

                var left = leftNode?.Text ?? "";
                var right = rightNode?.Text ?? "";
                var leftUnderscore = left.StartsWith("_", StringComparison.Ordinal);
                var rightUnderscore = right.StartsWith("_", StringComparison.Ordinal);
                if (leftUnderscore != rightUnderscore) return leftUnderscore ? -1 : 1;
                return StringComparer.CurrentCultureIgnoreCase.Compare(left, right);
            }
        }
    }
}
