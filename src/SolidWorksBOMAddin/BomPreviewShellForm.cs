using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using BomCore;
using SolidWorks.Interop.sldworks;

namespace SolidWorksBOMAddin;

internal sealed class BomPreviewShellForm : Form
{
    private static readonly Color ShellBackColor = Color.FromArgb(238, 241, 235);
    private static readonly Color ShellSurfaceColor = Color.FromArgb(252, 250, 244);
    private static readonly Color ShellHeaderColor = Color.FromArgb(31, 57, 53);
    private static readonly Color ShellAccentColor = Color.FromArgb(193, 122, 61);
    private static readonly Color ShellAccentDarkColor = Color.FromArgb(139, 82, 45);
    private static readonly Color ShellGridLineColor = Color.FromArgb(214, 207, 190);
    private static readonly Color ShellTextColor = Color.FromArgb(30, 34, 32);
    private static readonly Color ShellMutedTextColor = Color.FromArgb(89, 92, 86);

    private readonly BomPipeAddin _addin;
    private readonly ProfileStore _profileStore = new();
    private readonly PropertyDiscoveryService _propertyDiscoveryService = new();
    private readonly BomGenerator _bomGenerator = new();
    private readonly BindingList<PipeColumnMappingRow> _pipeColumns = [];
    private readonly Dictionary<string, BindingList<SectionColumnMappingRow>> _sectionColumnsBySection = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DataGridView> _sectionColumnGrids = new(StringComparer.OrdinalIgnoreCase);
    private readonly BindingList<SectionRuleMappingRow> _sectionRules = [];
    private readonly BindingList<AccessoryMappingRow> _accessoryRules = [];
    private readonly Dictionary<PropertyScope, DataGridView> _selectedPropertyGrids;
    private readonly DataGridView _pipeColumnsGrid;
    private readonly TabControl _sectionTabs;
    private readonly DataGridView _sectionRulesGrid;
    private readonly DataGridView _accessoryRulesGrid;
    private readonly DataGridView _previewGrid;
    private readonly DataGridView _diagnosticsGrid;
    private readonly DataGridView _mappingPreviewGrid;
    private readonly DataGridView _mappingDiagnosticsGrid;
    private readonly ListBox _discoveredPropertiesList;
    private readonly TabControl _selectedPropertiesTabs;
    private readonly Label _summaryLabel;
    private readonly Label _profileLabel;
    private readonly Label _statusLabel;
    private readonly TabControl _tabControl;

    private IReadOnlyList<ComponentRecord> _scannedComponents = [];
    private IReadOnlyList<BomDiagnostic> _profileDiagnostics = [];
    private BomProfile _currentProfile = new();
    private string? _assemblyPath;
    private string? _assemblyDisplayName;
    private string? _profileSourcePath;
    private string? _externalSettingsPath;
    private int _componentsScanned;
    private int? _componentsSkipped;

    public BomPreviewShellForm(BomPipeAddin addin)
    {
        _addin = addin ?? throw new ArgumentNullException(nameof(addin));

        Text = "BOMPipe Mapping Workspace";
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;
        BackColor = ShellBackColor;
        Font = new Font("Segoe UI", 9F);
        Width = 1280;
        Height = 860;
        MinimumSize = new Size(960, 640);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(12),
            BackColor = ShellBackColor,
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var commandPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = ShellBackColor,
        };

        commandPanel.Controls.Add(CreateActionButton("Read Selected Part", (_, _) => ReadSelectedPartProperties()));
        commandPanel.Controls.Add(CreateActionButton("Scan Assembly", (_, _) => ScanActiveAssembly()));
        commandPanel.Controls.Add(CreateActionButton("Mapping", (_, _) => ShowMappingTab()));
        commandPanel.Controls.Add(CreateActionButton("Refresh Preview", (_, _) => RefreshMappingPreview()));
        commandPanel.Controls.Add(CreateActionButton("Generate Preview", (_, _) => GenerateBomPreview()));
        commandPanel.Controls.Add(CreateActionButton("Export CSV", (_, _) => ExportBom("csv")));
        commandPanel.Controls.Add(CreateActionButton("Export Excel", (_, _) => ExportBom("xlsx")));
        commandPanel.Controls.Add(CreateActionButton("Import Settings", (_, _) => ImportSettings()));
        commandPanel.Controls.Add(CreateActionButton("Export Settings", (_, _) => ExportSettings()));
        commandPanel.Controls.Add(CreateActionButton("Set Settings File", (_, _) => SetSettingsFile()));
        commandPanel.Controls.Add(CreateActionButton("Save Mapping Profile", (_, _) => SaveMapping(), accent: true));

        var infoPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            AutoSize = true,
            Padding = new Padding(12),
            Margin = new Padding(0, 0, 0, 10),
            BackColor = ShellHeaderColor,
        };

        _summaryLabel = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font(Font, FontStyle.Bold),
            Text = "Start: select a component in SolidWorks, read its properties, scan the assembly, map columns, refresh preview, then save the mapping profile.",
        };
        _profileLabel = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ForeColor = Color.FromArgb(224, 230, 219),
            Text = "Rows are property-driven. Component names, file paths, and configuration names are diagnostics only.",
        };
        infoPanel.Controls.Add(_summaryLabel, 0, 0);
        infoPanel.Controls.Add(_profileLabel, 0, 1);

        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill,
        };

        _selectedPropertyGrids = CreateSelectedPropertyGrids();
        _selectedPropertiesTabs = CreateSelectedPropertiesTabs();
        _pipeColumnsGrid = CreatePipeColumnsGrid();
        _accessoryRulesGrid = CreateAccessoryRulesGrid();
        _sectionTabs = CreateSectionTabs();
        _sectionRulesGrid = CreateSectionRulesGrid();
        _previewGrid = CreateReadOnlyGrid();
        _diagnosticsGrid = CreateReadOnlyGrid();
        _mappingPreviewGrid = CreateReadOnlyGrid();
        _mappingDiagnosticsGrid = CreateReadOnlyGrid();
        _discoveredPropertiesList = new ListBox
        {
            Dock = DockStyle.Fill,
            IntegralHeight = false,
            BackColor = ShellSurfaceColor,
            ForeColor = ShellTextColor,
            BorderStyle = BorderStyle.None,
        };

        _tabControl.TabPages.Add(CreateSelectedPropertiesTab());
        _tabControl.TabPages.Add(CreateMappingTab());
        _tabControl.TabPages.Add(CreatePreviewTab());
        _tabControl.TabPages.Add(CreateDiagnosticsTab());

        _statusLabel = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Padding = new Padding(0, 6, 0, 0),
            ForeColor = ShellMutedTextColor,
            BackColor = ShellBackColor,
            Text = "Ready.",
        };

        root.Controls.Add(commandPanel, 0, 0);
        root.Controls.Add(infoPanel, 0, 1);
        root.Controls.Add(_tabControl, 0, 2);
        root.Controls.Add(_statusLabel, 0, 3);
        Controls.Add(root);

        _externalSettingsPath = LoadConfiguredSettingsPath();
        LoadDefaultProfileContext();
    }

    private static Button CreateActionButton(string text, EventHandler onClick, bool accent = false)
    {
        var button = new Button
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 0, 8, 8),
            Padding = new Padding(10, 6, 10, 6),
            Text = text,
            UseVisualStyleBackColor = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = accent ? ShellAccentColor : ShellSurfaceColor,
            ForeColor = accent ? Color.White : ShellTextColor,
        };
        button.FlatAppearance.BorderColor = accent ? ShellAccentDarkColor : ShellGridLineColor;
        button.FlatAppearance.MouseOverBackColor = accent ? ShellAccentDarkColor : Color.FromArgb(247, 241, 226);
        button.Click += onClick;
        return button;
    }

    private static DataGridView CreateReadOnlyGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        };
        StyleGrid(grid);
        return grid;
    }

    private static void StyleGrid(DataGridView grid)
    {
        grid.BackgroundColor = ShellSurfaceColor;
        grid.BorderStyle = BorderStyle.None;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
        grid.EnableHeadersVisualStyles = false;
        grid.GridColor = ShellGridLineColor;
        grid.DefaultCellStyle.BackColor = ShellSurfaceColor;
        grid.DefaultCellStyle.ForeColor = ShellTextColor;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(58, 95, 89);
        grid.DefaultCellStyle.SelectionForeColor = Color.White;
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(247, 244, 235);
        grid.ColumnHeadersDefaultCellStyle.BackColor = ShellHeaderColor;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        grid.RowTemplate.Height = 26;
    }

    private DataGridView CreatePipeColumnsGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            DataSource = _pipeColumns,
        };
        StyleGrid(grid);

        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.SourceProperty), "Source Property"));
        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.DisplayName), "Display Name"));
        grid.Columns.Add(CreateCheckboxColumn(nameof(PipeColumnMappingRow.Enabled), "Enabled"));
        grid.Columns.Add(CreateCheckboxColumn(nameof(PipeColumnMappingRow.GroupBy), "Group By"));
        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.Order), "Order", 80));
        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.Unit), "Unit", 90));
        return grid;
    }

    private TabControl CreateSectionTabs()
    {
        var tabControl = new TabControl
        {
            Dock = DockStyle.Fill,
        };

        foreach (var section in KnownBomSections.ConfigurableSections)
        {
            var rows = new BindingList<SectionColumnMappingRow>();
            _sectionColumnsBySection[section] = rows;

            var grid = CreateSectionColumnsGrid(rows);
            _sectionColumnGrids[section] = grid;

            var page = new TabPage(section)
            {
                Name = $"{section}SectionTab",
                BackColor = ShellSurfaceColor,
            };
            page.Controls.Add(grid);
            tabControl.TabPages.Add(page);
        }

        var accessoryPage = new TabPage(KnownBomSections.OtherAccessories)
        {
            Name = "OtherAccessoriesSectionTab",
            BackColor = ShellSurfaceColor,
        };
        accessoryPage.Controls.Add(_accessoryRulesGrid);
        tabControl.TabPages.Add(accessoryPage);

        return tabControl;
    }

    private DataGridView CreateSectionColumnsGrid(BindingList<SectionColumnMappingRow> rows)
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            DataSource = rows,
        };
        StyleGrid(grid);

        grid.Columns.Add(CreateTextColumn(nameof(SectionColumnMappingRow.SourceProperty), "Source Property"));
        grid.Columns.Add(CreateTextColumn(nameof(SectionColumnMappingRow.DisplayName), "Output Header"));
        grid.Columns.Add(CreateCheckboxColumn(nameof(SectionColumnMappingRow.Enabled), "Enabled"));
        grid.Columns.Add(CreateCheckboxColumn(nameof(SectionColumnMappingRow.GroupBy), "Group"));
        grid.Columns.Add(CreateTextColumn(nameof(SectionColumnMappingRow.Order), "Order", 70));
        grid.Columns.Add(CreateTextColumn(nameof(SectionColumnMappingRow.Unit), "Unit", 70));
        return grid;
    }

    private DataGridView CreateSectionRulesGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            DataSource = _sectionRules,
        };
        StyleGrid(grid);

        grid.Columns.Add(CreateTextColumn(nameof(SectionRuleMappingRow.SourceProperty), "Family Property", 120));
        grid.Columns.Add(CreateTextColumn(nameof(SectionRuleMappingRow.MatchValue), "Match Value"));
        grid.Columns.Add(CreateTextColumn(nameof(SectionRuleMappingRow.Section), "Output Section"));
        return grid;
    }

    private DataGridView CreateAccessoryRulesGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            DataSource = _accessoryRules,
        };
        StyleGrid(grid);

        grid.Columns.Add(CreateTextColumn(nameof(AccessoryMappingRow.SourceProperty), "Quantity Property"));
        grid.Columns.Add(CreateTextColumn(nameof(AccessoryMappingRow.DisplayName), "Accessory Name"));
        return grid;
    }

    private static DataGridViewTextBoxColumn CreateTextColumn(string dataPropertyName, string headerText, int minimumWidth = 120)
    {
        return new DataGridViewTextBoxColumn
        {
            DataPropertyName = dataPropertyName,
            HeaderText = headerText,
            MinimumWidth = minimumWidth,
            SortMode = DataGridViewColumnSortMode.NotSortable,
        };
    }

    private static DataGridViewCheckBoxColumn CreateCheckboxColumn(string dataPropertyName, string headerText)
    {
        return new DataGridViewCheckBoxColumn
        {
            DataPropertyName = dataPropertyName,
            HeaderText = headerText,
            MinimumWidth = 80,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
        };
    }

    private TabPage CreateSelectedPropertiesTab()
    {
        var page = new TabPage("Selected Properties")
        {
            Name = "SelectedPropertiesTab",
            BackColor = ShellBackColor,
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(8),
            BackColor = ShellBackColor,
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var commands = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
            BackColor = ShellBackColor,
        };
        commands.Controls.Add(CreateActionButton("Read Selected Part", (_, _) => ReadSelectedPartProperties(), accent: true));
        commands.Controls.Add(CreateActionButton("Add Property To Active Family", (_, _) => AddSelectedPartPropertyToMapping()));
        commands.Controls.Add(CreateActionButton("Scan Assembly", (_, _) => ScanActiveAssembly()));

        var hint = new Label
        {
            AutoSize = true,
            ForeColor = ShellMutedTextColor,
            Padding = new Padding(8, 7, 0, 0),
            Text = "Select part in SolidWorks, read properties here, then add highlighted property to active family tab in BOM Mapping.",
        };
        commands.Controls.Add(hint);

        layout.Controls.Add(commands, 0, 0);
        layout.Controls.Add(_selectedPropertiesTabs, 0, 1);

        page.Controls.Add(layout);
        return page;
    }

    private TabControl CreateSelectedPropertiesTabs()
    {
        var tabs = new TabControl
        {
            Dock = DockStyle.Fill,
        };

        foreach (var scope in GetPropertyScopeDisplayOrder())
        {
            var scopePage = new TabPage(GetPropertyScopeTabText(scope))
            {
                Name = $"{scope}PropertiesTab",
            };
            scopePage.Controls.Add(_selectedPropertyGrids[scope]);
            tabs.TabPages.Add(scopePage);
        }

        return tabs;
    }

    private TabPage CreateMappingTab()
    {
        var page = new TabPage("BOM Mapping")
        {
            Name = "MappingTab",
            BackColor = ShellBackColor,
        };
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Padding = new Padding(4),
            BackColor = ShellBackColor,
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260F));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58F));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42F));

        var discoveredGroup = new GroupBox
        {
            Dock = DockStyle.Fill,
            Text = "Discovered Properties",
            Padding = new Padding(8),
            BackColor = ShellSurfaceColor,
            ForeColor = ShellTextColor,
        };
        discoveredGroup.Controls.Add(_discoveredPropertiesList);

        var center = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            BackColor = ShellBackColor,
        };
        center.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        center.RowStyles.Add(new RowStyle(SizeType.Absolute, 130F));
        center.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var sectionGroup = new GroupBox
        {
            Dock = DockStyle.Fill,
            Text = "Section Columns",
            Padding = new Padding(8),
            BackColor = ShellSurfaceColor,
            ForeColor = ShellTextColor,
        };
        sectionGroup.Controls.Add(_sectionTabs);

        var rulesGroup = new GroupBox
        {
            Dock = DockStyle.Fill,
            Text = "Primary Family Rules",
            Padding = new Padding(8),
            BackColor = ShellSurfaceColor,
            ForeColor = ShellTextColor,
        };
        rulesGroup.Controls.Add(_sectionRulesGrid);

        var centerCommands = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
        };
        centerCommands.BackColor = ShellBackColor;
        centerCommands.Controls.Add(CreateActionButton("Add Discovered Column", (_, _) => AddSelectedDiscoveredProperty()));
        centerCommands.Controls.Add(CreateActionButton("Remove Selected Column", (_, _) => RemoveSelectedSectionColumn()));
        centerCommands.Controls.Add(CreateActionButton("Refresh Preview", (_, _) => RefreshMappingPreview(), accent: true));
        centerCommands.Controls.Add(CreateActionButton("Import Settings", (_, _) => ImportSettings()));
        centerCommands.Controls.Add(CreateActionButton("Export Settings", (_, _) => ExportSettings()));
        centerCommands.Controls.Add(CreateActionButton("Save Mapping Profile", (_, _) => SaveMapping()));

        center.Controls.Add(sectionGroup, 0, 0);
        center.Controls.Add(rulesGroup, 0, 1);
        center.Controls.Add(centerCommands, 0, 2);

        var right = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = ShellBackColor,
        };
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 60F));
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 40F));

        var previewGroup = new GroupBox
        {
            Dock = DockStyle.Fill,
            Text = "Live Preview",
            Padding = new Padding(8),
            BackColor = ShellSurfaceColor,
            ForeColor = ShellTextColor,
        };
        previewGroup.Controls.Add(_mappingPreviewGrid);

        var diagnosticsGroup = new GroupBox
        {
            Dock = DockStyle.Fill,
            Text = "Diagnostics",
            Padding = new Padding(8),
            BackColor = ShellSurfaceColor,
            ForeColor = ShellTextColor,
        };
        diagnosticsGroup.Controls.Add(_mappingDiagnosticsGrid);

        right.Controls.Add(previewGroup, 0, 0);
        right.Controls.Add(diagnosticsGroup, 0, 1);

        layout.Controls.Add(discoveredGroup, 0, 0);
        layout.Controls.Add(center, 1, 0);
        layout.Controls.Add(right, 2, 0);

        page.Controls.Add(layout);
        return page;
    }

    private TabPage CreatePreviewTab()
    {
        var page = new TabPage("BOM Preview")
        {
            Name = "PreviewTab",
            BackColor = ShellBackColor,
        };
        page.Controls.Add(_previewGrid);
        return page;
    }

    private TabPage CreateDiagnosticsTab()
    {
        var page = new TabPage("Diagnostics")
        {
            Name = "DiagnosticsTab",
            BackColor = ShellBackColor,
        };
        page.Controls.Add(_diagnosticsGrid);
        return page;
    }

    private void ReadSelectedPartProperties()
    {
        try
        {
            var properties = _addin.SelectedComponentPropertyReader?.ReadSelectedComponentProperties()
                ?? throw new InvalidOperationException("Selected component property reader is unavailable.");

            foreach (var scope in GetPropertyScopeDisplayOrder())
            {
                var table = CreateSelectedPropertiesTable();
                foreach (var property in properties
                             .Where(property => property.Scope == scope)
                             .OrderBy(property => property.Name, StringComparer.OrdinalIgnoreCase)
                             .ThenBy(property => property.Source, StringComparer.OrdinalIgnoreCase))
                {
                    table.Rows.Add(
                        property.Name,
                        property.EffectiveValue,
                        property.RawValue ?? string.Empty,
                        property.Source ?? string.Empty);
                }

                _selectedPropertyGrids[scope].DataSource = table;
            }

            _tabControl.SelectedTab = _tabControl.TabPages["SelectedPropertiesTab"] ?? _tabControl.TabPages[0];
            SetStatus(properties.Count == 0
                ? "No selected part properties were returned."
                : $"Loaded {properties.Count} selected part propert{(properties.Count == 1 ? "y" : "ies")}.");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void ScanActiveAssembly()
    {
        try
        {
            var application = _addin.RequireApplication();
            var activeDocument = application.IActiveDoc2
                ?? throw new InvalidOperationException("No active SolidWorks document is open.");
            if (activeDocument is not IAssemblyDoc)
            {
                throw new InvalidOperationException("The active SolidWorks document is not an assembly.");
            }

            var assemblyReadResult = _addin.AssemblyReader?.ReadActiveAssembly()
                ?? throw new InvalidOperationException("Assembly reader is unavailable.");

            _scannedComponents = assemblyReadResult.Components;
            _componentsScanned = assemblyReadResult.ComponentsScanned;
            _componentsSkipped = assemblyReadResult.ComponentsSkipped;
            _assemblyPath = string.IsNullOrWhiteSpace(assemblyReadResult.AssemblyPath)
                ? (string.IsNullOrWhiteSpace(activeDocument.GetPathName()) ? null : activeDocument.GetPathName())
                : assemblyReadResult.AssemblyPath;
            _assemblyDisplayName = !string.IsNullOrWhiteSpace(_assemblyPath)
                ? Path.GetFileName(_assemblyPath)
                : activeDocument.GetTitle();

            var profileLoadResult = LoadEffectiveProfile(_assemblyPath);
            _currentProfile = profileLoadResult.Profile;
            _profileSourcePath = profileLoadResult.SourcePath;
            _profileDiagnostics = profileLoadResult.Diagnostics
                .Concat(assemblyReadResult.Diagnostics)
                .ToList();

            var discovery = _propertyDiscoveryService.DiscoverFromComponents(_scannedComponents);
            BindProfile(_currentProfile, discovery);
            BindDiagnostics(_profileDiagnostics);
            BindMappingDiagnostics(_profileDiagnostics);
            UpdateSummary();

            _tabControl.SelectedTab = _tabControl.TabPages["MappingTab"] ?? _tabControl.TabPages[1];
            SetStatus($"Scanned {_componentsScanned} component(s) from {_assemblyDisplayName ?? "the active assembly"}.");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void ShowMappingTab()
    {
        _tabControl.SelectedTab = _tabControl.TabPages["MappingTab"] ?? _tabControl.TabPages[1];
        SetStatus("Review or edit the BOM mapping, then generate a preview or export.");
    }

    private void GenerateBomPreview()
    {
        try
        {
            if (_scannedComponents.Count == 0)
            {
                ScanActiveAssembly();
                if (_scannedComponents.Count == 0)
                {
                    return;
                }
            }

            var profile = BuildProfileFromEditor();
            var result = _bomGenerator.Generate(_scannedComponents, profile);
            var diagnostics = _profileDiagnostics
                .Concat(result.Diagnostics)
                .ToList();

            BindPreview(result);
            BindDiagnostics(diagnostics);
            BindMappingPreview(result);
            BindMappingDiagnostics(diagnostics);
            _tabControl.SelectedTab = _tabControl.TabPages["PreviewTab"] ?? _tabControl.TabPages[2];
            SetStatus($"Generated {result.Rows.Count} BOM row(s) using {profile.GetEffectiveSectionColumnProfiles().Sum(section => section.Columns.Count)} section column rule(s) and {profile.AccessoryRules.Count} accessory rule(s).");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void ExportBom(string format)
    {
        try
        {
            if (_scannedComponents.Count == 0)
            {
                ScanActiveAssembly();
                if (_scannedComponents.Count == 0)
                {
                    return;
                }
            }

            var profile = BuildProfileFromEditor();
            var result = _bomGenerator.Generate(_scannedComponents, profile);
            var diagnostics = _profileDiagnostics.Concat(result.Diagnostics).ToList();

            using var dialog = new SaveFileDialog
            {
                AddExtension = true,
                DefaultExt = format,
                Filter = string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase)
                    ? "Excel Workbook (*.xlsx)|*.xlsx"
                    : "CSV (*.csv)|*.csv",
                FileName = BuildDefaultExportFileName(format),
                InitialDirectory = ResolveInitialExportDirectory(),
                OverwritePrompt = true,
                Title = string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase) ? "Export BOM Preview to Excel" : "Export BOM Preview to CSV",
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                SetStatus("Export cancelled.");
                return;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(dialog.FileName)!);
            using var stream = File.Create(dialog.FileName);
            CreateExporter(format).Export(
                new BomResult
                {
                    Rows = result.Rows,
                    Diagnostics = diagnostics,
                },
                stream);

            BindPreview(result);
            BindDiagnostics(diagnostics);
            BindMappingPreview(result);
            BindMappingDiagnostics(diagnostics);
            _tabControl.SelectedTab = _tabControl.TabPages["PreviewTab"] ?? _tabControl.TabPages[2];
            SetStatus($"Exported {result.Rows.Count} BOM row(s) to {dialog.FileName}.");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void SaveMapping()
    {
        try
        {
            var profile = BuildProfileFromEditor();
            var targetPath = ResolveProfileSavePath();
            _profileStore.SaveToPath(profile, targetPath);
            _profileSourcePath = targetPath;
            _profileDiagnostics = BomProfileSerializer.Validate(profile);
            BindDiagnostics(_profileDiagnostics);
            BindMappingDiagnostics(_profileDiagnostics);
            UpdateSummary();
            SetStatus($"Saved BOM mapping to {targetPath}.");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void LoadDefaultProfileContext()
    {
        var profileLoadResult = LoadEffectiveProfile(assemblyPath: null);
        _currentProfile = profileLoadResult.Profile;
        _profileSourcePath = profileLoadResult.SourcePath;
        _profileDiagnostics = profileLoadResult.Diagnostics;
        BindProfile(_currentProfile, discovery: null);
        BindDiagnostics(_profileDiagnostics);
        BindMappingDiagnostics(_profileDiagnostics);
        UpdateSummary();
    }

    private ProfileLoadResult LoadEffectiveProfile(string? assemblyPath)
    {
        var options = BuildProfileStoreOptions();
        if (!string.IsNullOrWhiteSpace(_externalSettingsPath) && File.Exists(_externalSettingsPath))
        {
            try
            {
                var externalProfile = _profileStore.LoadFromPath(_externalSettingsPath);
                var diagnostics = BomProfileSerializer.Validate(externalProfile);
                if (diagnostics.Any(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error))
                {
                    return LoadBuiltInProfileWithDiagnostics(
                        options.DefaultProfilePath,
                        diagnostics.Prepend(new BomDiagnostic
                        {
                            Severity = DiagnosticSeverity.Warning,
                            Code = "external-settings-invalid",
                            Message = $"Settings file '{_externalSettingsPath}' is invalid. Falling back to the built-in default profile.",
                        }));
                }

                return new ProfileLoadResult
                {
                    Profile = externalProfile,
                    SourcePath = _externalSettingsPath,
                    Diagnostics = diagnostics,
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
            {
                return LoadBuiltInProfileWithDiagnostics(
                    options.DefaultProfilePath,
                    [
                        new BomDiagnostic
                        {
                            Severity = DiagnosticSeverity.Warning,
                            Code = "external-settings-load-failed",
                            Message = $"Settings file '{_externalSettingsPath}' could not be loaded. Falling back to the built-in default profile.",
                        },
                    ]);
            }
        }

        if (!string.IsNullOrWhiteSpace(assemblyPath))
        {
            return _profileStore.LoadEffectiveProfile(assemblyPath, options);
        }

        var profile = _profileStore.LoadFromPath(options.DefaultProfilePath);
        return new ProfileLoadResult
        {
            Profile = profile,
            SourcePath = options.DefaultProfilePath,
            Diagnostics = BomProfileSerializer.Validate(profile),
        };
    }

    private ProfileStoreOptions BuildProfileStoreOptions()
    {
        return new ProfileStoreOptions
        {
            DefaultProfilePath = Path.Combine(GetAddinDirectory(), "profiles", "default.pipebom.json"),
            UserProfileDirectory = Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                "AFCA",
                "SolidWorksBOMAddin",
                "profiles"),
            CompanyProfileDirectory = Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData),
                "AFCA",
                "SolidWorksBOMAddin",
                "profiles"),
        };
    }

    private ProfileLoadResult LoadBuiltInProfileWithDiagnostics(string defaultProfilePath, IEnumerable<BomDiagnostic> diagnostics)
    {
        var profile = _profileStore.LoadFromPath(defaultProfilePath);
        var mergedDiagnostics = diagnostics
            .Concat(BomProfileSerializer.Validate(profile))
            .ToList();

        return new ProfileLoadResult
        {
            Profile = profile,
            SourcePath = defaultProfilePath,
            Diagnostics = mergedDiagnostics,
        };
    }

    private static string GetAddinDirectory()
    {
        return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            ?? AppContext.BaseDirectory;
    }

    private void BindProfile(BomProfile profile, PropertyDiscoveryResult? discovery)
    {
        _pipeColumns.Clear();
        foreach (var column in profile.GetSectionColumns(KnownBomSections.Pipes).OrderBy(column => column.Order))
        {
            _pipeColumns.Add(new PipeColumnMappingRow
            {
                SourceProperty = column.SourceProperty,
                DisplayName = column.DisplayName,
                Enabled = column.Enabled,
                GroupBy = column.GroupBy,
                Order = column.Order,
                Unit = column.Unit ?? string.Empty,
            });
        }

        foreach (var section in KnownBomSections.ConfigurableSections)
        {
            if (!_sectionColumnsBySection.TryGetValue(section, out var rows))
            {
                continue;
            }

            rows.Clear();
            foreach (var column in profile.GetSectionColumns(section).OrderBy(column => column.Order))
            {
                rows.Add(new SectionColumnMappingRow
                {
                    SourceProperty = column.SourceProperty,
                    DisplayName = column.DisplayName,
                    Enabled = column.Enabled,
                    GroupBy = column.GroupBy,
                    Order = column.Order,
                    Unit = column.Unit ?? string.Empty,
                });
            }
        }

        _sectionRules.Clear();
        foreach (var rule in profile.GetEffectiveSectionRules())
        {
            _sectionRules.Add(new SectionRuleMappingRow
            {
                SourceProperty = string.IsNullOrWhiteSpace(rule.SourceProperty) ? KnownPropertyNames.PrimaryFamily : rule.SourceProperty,
                MatchValue = rule.MatchValue,
                Section = KnownBomSections.NormalizeConfigurableSection(rule.Section),
            });
        }

        _accessoryRules.Clear();
        foreach (var accessoryRule in profile.AccessoryRules)
        {
            _accessoryRules.Add(new AccessoryMappingRow
            {
                SourceProperty = accessoryRule.SourceProperty,
                DisplayName = accessoryRule.DisplayName,
                BomSection = KnownBomSections.NormalizeAccessorySection(accessoryRule.BomSection),
            });
        }

        _discoveredPropertiesList.BeginUpdate();
        _discoveredPropertiesList.Items.Clear();
        if (discovery is not null)
        {
            foreach (var propertyName in BuildDiscoveredPropertyItems(discovery.DiscoveredProperties))
            {
                _discoveredPropertiesList.Items.Add(propertyName);
            }
        }
        _discoveredPropertiesList.EndUpdate();
        UpdateSummary();
        TryRefreshMappingPreview();
    }

    private void LoadSettingsFile(string path, bool persistAsActiveSettings)
    {
        var profile = _profileStore.LoadFromPath(path);
        var diagnostics = BomProfileSerializer.Validate(profile);
        if (diagnostics.Any(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error))
        {
            throw new InvalidOperationException($"Settings file '{path}' contains invalid BOM profile settings.");
        }

        _currentProfile = profile;
        _profileSourcePath = path;
        _profileDiagnostics = diagnostics;
        if (persistAsActiveSettings)
        {
            ActivateExternalSettingsPath(path);
        }

        var discovery = _scannedComponents.Count == 0
            ? null
            : _propertyDiscoveryService.DiscoverFromComponents(_scannedComponents);
        BindProfile(_currentProfile, discovery);
        BindDiagnostics(_profileDiagnostics);
        BindMappingDiagnostics(_profileDiagnostics);
        UpdateSummary();
        TryRefreshMappingPreview();
    }

    private void ActivateExternalSettingsPath(string path)
    {
        _externalSettingsPath = path;
        SaveConfiguredSettingsPath(path);
    }

    private void BindPreview(BomResult result)
    {
        _previewGrid.DataSource = BuildPreviewTable(result);
    }

    private void BindMappingPreview(BomResult result)
    {
        _mappingPreviewGrid.DataSource = BuildPreviewTable(result);
    }

    private static DataTable BuildPreviewTable(BomResult result)
    {
        var headers = CollectHeaders(result.Rows);
        var table = new DataTable();
        table.Columns.Add("Section");
        table.Columns.Add("Row Type");
        foreach (var header in headers)
        {
            table.Columns.Add(header);
        }

        table.Columns.Add("Quantity");

        foreach (var row in result.Rows)
        {
            var item = table.NewRow();
            item["Section"] = row.Section;
            item["Row Type"] = row.RowType.ToString();
            foreach (var header in headers)
            {
                item[header] = row.Values.TryGetValue(header, out var value) ? value : string.Empty;
            }

            item["Quantity"] = row.Quantity.ToString(CultureInfo.InvariantCulture);
            table.Rows.Add(item);
        }

        return table;
    }

    private void BindDiagnostics(IEnumerable<BomDiagnostic> diagnostics)
    {
        _diagnosticsGrid.DataSource = BuildDiagnosticsTable(diagnostics);
    }

    private void BindMappingDiagnostics(IEnumerable<BomDiagnostic> diagnostics)
    {
        _mappingDiagnosticsGrid.DataSource = BuildDiagnosticsTable(diagnostics);
    }

    private static DataTable BuildDiagnosticsTable(IEnumerable<BomDiagnostic> diagnostics)
    {
        var table = new DataTable();
        table.Columns.Add("Severity");
        table.Columns.Add("Code");
        table.Columns.Add("Message");
        table.Columns.Add("Component");
        table.Columns.Add("Property");

        foreach (var diagnostic in diagnostics)
        {
            table.Rows.Add(
                diagnostic.Severity.ToString(),
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.ComponentId ?? string.Empty,
                diagnostic.PropertyName ?? string.Empty);
        }

        return table;
    }

    private void UpdateSummary()
    {
        var assemblyText = string.IsNullOrWhiteSpace(_assemblyDisplayName) ? "No assembly scanned" : _assemblyDisplayName;
        var componentSummary = _componentsScanned > 0
            ? $"Scanned: {_componentsScanned} | Loaded: {_scannedComponents.Count} | Skipped: {_componentsSkipped ?? 0}"
            : $"Loaded: {_scannedComponents.Count}";
        _summaryLabel.Text = $"Assembly: {assemblyText} | {componentSummary}";
        var settingsText = string.IsNullOrWhiteSpace(_externalSettingsPath)
            ? "Local/default lookup"
            : _externalSettingsPath;
        _profileLabel.Text = $"Profile: {_profileSourcePath ?? "Built-in default"} | Settings file: {settingsText} | Sections: {string.Join(", ", KnownBomSections.ConfigurableSections)}";
    }

    private IReadOnlyList<string> BuildDiscoveredPropertyItems(IEnumerable<string> propertyNames)
    {
        var propertyScopeByName = _scannedComponents
            .SelectMany(component => component.Properties.Values)
            .GroupBy(property => property.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => group.Select(property => property.Scope).FirstOrDefault(),
                StringComparer.OrdinalIgnoreCase);

        return propertyNames
            .OrderBy(propertyName => propertyScopeByName.TryGetValue(propertyName, out var scope) ? scope : PropertyScope.Unknown)
            .ThenBy(propertyName => propertyName, StringComparer.OrdinalIgnoreCase)
            .Select(propertyName =>
            {
                var scope = propertyScopeByName.TryGetValue(propertyName, out var knownScope) ? knownScope : PropertyScope.Unknown;
                return $"{scope}: {propertyName}";
            })
            .ToList();
    }

    private void SetStatus(string message)
    {
        _statusLabel.Text = message;
    }

    private void ShowError(string message)
    {
        SetStatus(message);
        MessageBox.Show(this, message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    private BomProfile BuildProfileFromEditor()
    {
        _pipeColumnsGrid.EndEdit();
        _sectionRulesGrid.EndEdit();
        _accessoryRulesGrid.EndEdit();
        foreach (var grid in _sectionColumnGrids.Values)
        {
            grid.EndEdit();
        }

        var sectionProfiles = KnownBomSections.ConfigurableSections
            .Select(section => new BomSectionColumnProfile
            {
                Section = section,
                Columns = BuildSectionColumns(section),
            })
            .ToList();

        var pipeColumns = sectionProfiles
            .First(profile => string.Equals(profile.Section, KnownBomSections.Pipes, StringComparison.OrdinalIgnoreCase))
            .Columns;

        var sectionRules = _sectionRules
            .Where(row => !string.IsNullOrWhiteSpace(row.MatchValue) && !string.IsNullOrWhiteSpace(row.Section))
            .Select(row => new BomSectionRule
            {
                SourceProperty = string.IsNullOrWhiteSpace(row.SourceProperty) ? KnownPropertyNames.PrimaryFamily : row.SourceProperty.Trim(),
                MatchValue = row.MatchValue.Trim(),
                Section = KnownBomSections.NormalizeConfigurableSection(row.Section.Trim()),
            })
            .ToList();

        var accessoryRules = _accessoryRules
            .Where(row => !string.IsNullOrWhiteSpace(row.SourceProperty) && !string.IsNullOrWhiteSpace(row.DisplayName))
            .Select(row => new AccessoryRule
            {
                SourceProperty = row.SourceProperty.Trim(),
                DisplayName = row.DisplayName.Trim(),
                BomSection = KnownBomSections.NormalizeAccessorySection(row.BomSection),
            })
            .ToList();

        _currentProfile = _currentProfile with
        {
            PipeColumns = pipeColumns,
            SectionColumnProfiles = sectionProfiles,
            SectionRules = sectionRules,
            AccessoryRules = accessoryRules,
            ProfileName = string.IsNullOrWhiteSpace(_currentProfile.ProfileName) ? "AFCA Pipe BOM" : _currentProfile.ProfileName,
            Version = Math.Max(_currentProfile.Version, 2),
        };

        _profileDiagnostics = BomProfileSerializer.Validate(_currentProfile);
        return _currentProfile;
    }

    private IReadOnlyList<BomColumnRule> BuildSectionColumns(string section)
    {
        var orderFallback = 1;
        if (!_sectionColumnsBySection.TryGetValue(section, out var rows))
        {
            return KnownBomColumnProfiles.CreateDefaultSectionColumns(section);
        }

        return rows
            .Where(row => !string.IsNullOrWhiteSpace(row.SourceProperty) && !string.IsNullOrWhiteSpace(row.DisplayName))
            .Select(row => new BomColumnRule
            {
                SourceProperty = row.SourceProperty.Trim(),
                DisplayName = row.DisplayName.Trim(),
                Enabled = row.Enabled,
                GroupBy = row.GroupBy,
                Order = row.Order > 0 ? row.Order : orderFallback++,
                Unit = string.IsNullOrWhiteSpace(row.Unit) ? null : row.Unit.Trim(),
            })
            .OrderBy(column => column.Order)
            .ToList();
    }

    private void AddSelectedDiscoveredProperty()
    {
        if (_discoveredPropertiesList.SelectedItem is not string selectedItem)
        {
            SetStatus("Select a discovered property first.");
            return;
        }

        var propertyName = selectedItem.Contains(":", StringComparison.Ordinal)
            ? selectedItem[(selectedItem.IndexOf(":", StringComparison.Ordinal) + 1)..].Trim()
            : selectedItem.Trim();

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return;
        }

        AddPropertyToActiveSection(propertyName);
    }

    private void AddSelectedPartPropertyToMapping()
    {
        if (!TryGetSelectedPartPropertyName(out var propertyName))
        {
            SetStatus("Select a property row first, or read the selected SolidWorks part.");
            return;
        }

        AddPropertyToActiveSection(propertyName);
        _tabControl.SelectedTab = _tabControl.TabPages["MappingTab"] ?? _tabControl.TabPages[1];
    }

    private void AddPropertyToActiveSection(string propertyName)
    {
        var section = _sectionTabs.SelectedTab?.Text ?? KnownBomSections.Components;
        if (string.Equals(section, KnownBomSections.OtherAccessories, StringComparison.OrdinalIgnoreCase))
        {
            _accessoryRules.Add(new AccessoryMappingRow
            {
                SourceProperty = propertyName,
                DisplayName = propertyName,
                BomSection = KnownBomSections.OtherAccessories,
            });
            RefreshMappingPreview();
            SetStatus($"Added '{propertyName}' to {KnownBomSections.OtherAccessories}. Save mapping profile to persist.");
            return;
        }

        if (!_sectionColumnsBySection.TryGetValue(section, out var rows))
        {
            return;
        }

        rows.Add(new SectionColumnMappingRow
        {
            SourceProperty = propertyName,
            DisplayName = propertyName,
            Enabled = true,
            GroupBy = true,
            Order = rows.Count == 0 ? 1 : rows.Max(row => row.Order) + 1,
        });

        RefreshMappingPreview();
        SetStatus($"Added '{propertyName}' to {section} columns. Save mapping profile to persist.");
    }

    private void RemoveSelectedSectionColumn()
    {
        var section = _sectionTabs.SelectedTab?.Text ?? KnownBomSections.Components;
        if (string.Equals(section, KnownBomSections.OtherAccessories, StringComparison.OrdinalIgnoreCase))
        {
            RemoveSelectedAccessoryRule();
            return;
        }

        if (!_sectionColumnsBySection.TryGetValue(section, out var rows)
            || !_sectionColumnGrids.TryGetValue(section, out var grid))
        {
            return;
        }

        var removed = 0;
        foreach (DataGridViewRow selectedRow in grid.SelectedRows.Cast<DataGridViewRow>().Where(row => !row.IsNewRow).ToList())
        {
            if (selectedRow.DataBoundItem is not SectionColumnMappingRow row)
            {
                continue;
            }

            rows.Remove(row);
            removed++;
        }

        if (removed == 0)
        {
            SetStatus("Select one or more columns in the active family tab before removing.");
            return;
        }

        RefreshMappingPreview();
        SetStatus($"Removed {removed} column{(removed == 1 ? string.Empty : "s")} from {section}. Save mapping profile to persist.");
    }

    private void RemoveSelectedAccessoryRule()
    {
        var removed = 0;
        foreach (DataGridViewRow selectedRow in _accessoryRulesGrid.SelectedRows.Cast<DataGridViewRow>().Where(row => !row.IsNewRow).ToList())
        {
            if (selectedRow.DataBoundItem is not AccessoryMappingRow row)
            {
                continue;
            }

            _accessoryRules.Remove(row);
            removed++;
        }

        if (removed == 0)
        {
            SetStatus($"Select one or more rules in {KnownBomSections.OtherAccessories} before removing.");
            return;
        }

        RefreshMappingPreview();
        SetStatus($"Removed {removed} accessory rule{(removed == 1 ? string.Empty : "s")}. Save mapping profile to persist.");
    }

    private bool TryGetSelectedPartPropertyName(out string propertyName)
    {
        propertyName = string.Empty;
        var grid = _selectedPropertiesTabs.SelectedTab?.Controls.OfType<DataGridView>().FirstOrDefault();
        var row = grid?.SelectedRows.Cast<DataGridViewRow>().FirstOrDefault(row => !row.IsNewRow)
            ?? grid?.CurrentRow;

        if (row is null || row.IsNewRow)
        {
            return false;
        }

        propertyName = Convert.ToString(row.Cells["Property"].Value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
        return !string.IsNullOrWhiteSpace(propertyName);
    }

    private void TryRefreshMappingPreview()
    {
        if (_scannedComponents.Count == 0)
        {
            return;
        }

        try
        {
            RefreshMappingPreview();
        }
        catch
        {
            // Editing grids can briefly contain invalid values; explicit preview/export will surface the error.
        }
    }

    private void ImportSettings()
    {
        try
        {
            using var dialog = new OpenFileDialog
            {
                AddExtension = true,
                CheckFileExists = true,
                DefaultExt = "json",
                Filter = BuildSettingsFileFilter(),
                InitialDirectory = ResolveSettingsDialogDirectory(),
                Title = "Import BOMPipe settings",
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                SetStatus("Settings import cancelled.");
                return;
            }

            LoadSettingsFile(dialog.FileName, persistAsActiveSettings: true);
            SetStatus($"Imported settings from {dialog.FileName}. Future saves use this settings file.");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void ExportSettings()
    {
        try
        {
            using var dialog = new SaveFileDialog
            {
                AddExtension = true,
                DefaultExt = "pipebom.json",
                Filter = BuildSettingsFileFilter(),
                FileName = BuildDefaultSettingsFileName(),
                InitialDirectory = ResolveSettingsDialogDirectory(),
                OverwritePrompt = true,
                Title = "Export BOMPipe settings",
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                SetStatus("Settings export cancelled.");
                return;
            }

            var profile = BuildProfileFromEditor();
            _profileStore.SaveToPath(profile, dialog.FileName);
            ActivateExternalSettingsPath(dialog.FileName);
            _currentProfile = profile;
            _profileDiagnostics = BomProfileSerializer.Validate(profile);
            _profileSourcePath = dialog.FileName;
            BindDiagnostics(_profileDiagnostics);
            BindMappingDiagnostics(_profileDiagnostics);
            UpdateSummary();
            SetStatus($"Exported settings to {dialog.FileName}. Future saves use this settings file.");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void SetSettingsFile()
    {
        try
        {
            using var dialog = new SaveFileDialog
            {
                AddExtension = true,
                DefaultExt = "pipebom.json",
                Filter = BuildSettingsFileFilter(),
                FileName = BuildDefaultSettingsFileName(),
                InitialDirectory = ResolveSettingsDialogDirectory(),
                OverwritePrompt = false,
                Title = "Choose BOMPipe settings file",
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                SetStatus("Settings file selection cancelled.");
                return;
            }

            ActivateExternalSettingsPath(dialog.FileName);
            if (File.Exists(dialog.FileName))
            {
                LoadSettingsFile(dialog.FileName, persistAsActiveSettings: true);
                SetStatus($"Using settings file {dialog.FileName}.");
                return;
            }

            _profileSourcePath = dialog.FileName;
            UpdateSummary();
            SetStatus($"Settings file set to {dialog.FileName}. Save mapping profile to create it.");
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private void RefreshMappingPreview()
    {
        if (_scannedComponents.Count == 0)
        {
            BindMappingPreview(new BomResult());
            BindMappingDiagnostics(_profileDiagnostics);
            SetStatus("Scan an assembly to populate the live preview.");
            return;
        }

        var profile = BuildProfileFromEditor();
        var result = _bomGenerator.Generate(_scannedComponents, profile);
        var diagnostics = _profileDiagnostics
            .Concat(result.Diagnostics)
            .ToList();

        BindMappingPreview(result);
        BindMappingDiagnostics(diagnostics);
        SetStatus($"Preview refreshed with {result.Rows.Count} BOM row(s).");
    }

    private string ResolveProfileSavePath()
    {
        if (!string.IsNullOrWhiteSpace(_externalSettingsPath))
        {
            return _externalSettingsPath;
        }

        if (!string.IsNullOrWhiteSpace(_assemblyPath))
        {
            return Path.Combine(Path.GetDirectoryName(_assemblyPath)!, BuildProfileStoreOptions().DefaultProfileFileName);
        }

        if (!string.IsNullOrWhiteSpace(_profileSourcePath)
            && !string.Equals(_profileSourcePath, BuildProfileStoreOptions().DefaultProfilePath, StringComparison.OrdinalIgnoreCase))
        {
            return _profileSourcePath;
        }

        return Path.Combine(
            BuildProfileStoreOptions().UserProfileDirectory!,
            BuildProfileStoreOptions().DefaultProfileFileName);
    }

    private string BuildDefaultExportFileName(string format)
    {
        var baseName = !string.IsNullOrWhiteSpace(_assemblyPath)
            ? Path.GetFileNameWithoutExtension(_assemblyPath)
            : (!string.IsNullOrWhiteSpace(_assemblyDisplayName) ? Path.GetFileNameWithoutExtension(_assemblyDisplayName) : "pipe-bom-preview");

        return $"{SanitizeFileName(baseName)}.bom.{format}";
    }

    private string ResolveInitialExportDirectory()
    {
        if (!string.IsNullOrWhiteSpace(_assemblyPath))
        {
            return Path.GetDirectoryName(_assemblyPath)!;
        }

        return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
    }

    private string ResolveSettingsDialogDirectory()
    {
        if (!string.IsNullOrWhiteSpace(_externalSettingsPath))
        {
            var directory = Path.GetDirectoryName(_externalSettingsPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                return directory;
            }
        }

        if (!string.IsNullOrWhiteSpace(_profileSourcePath))
        {
            var directory = Path.GetDirectoryName(_profileSourcePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                return directory;
            }
        }

        if (!string.IsNullOrWhiteSpace(_assemblyPath))
        {
            return Path.GetDirectoryName(_assemblyPath)!;
        }

        return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
    }

    private string BuildDefaultSettingsFileName()
    {
        if (!string.IsNullOrWhiteSpace(_externalSettingsPath))
        {
            return Path.GetFileName(_externalSettingsPath);
        }

        var baseName = !string.IsNullOrWhiteSpace(_assemblyPath)
            ? Path.GetFileNameWithoutExtension(_assemblyPath)
            : "bompipe-settings";

        return $"{SanitizeFileName(baseName)}.pipebom.json";
    }

    private static string BuildSettingsFileFilter()
    {
        return "BOMPipe Settings (*.pipebom.json)|*.pipebom.json|JSON (*.json)|*.json|All files (*.*)|*.*";
    }

    private static string? LoadConfiguredSettingsPath()
    {
        var path = GetSettingsPointerPath();
        if (!File.Exists(path))
        {
            return null;
        }

        var configuredPath = File.ReadAllText(path).Trim();
        return string.IsNullOrWhiteSpace(configuredPath) ? null : configuredPath;
    }

    private static void SaveConfiguredSettingsPath(string settingsPath)
    {
        var pointerPath = GetSettingsPointerPath();
        var directory = Path.GetDirectoryName(pointerPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(pointerPath, settingsPath);
    }

    private static string GetSettingsPointerPath()
    {
        return Path.Combine(
            System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
            "AFCA",
            "SolidWorksBOMAddin",
            "settings-file.path");
    }

    private static IBomExporter CreateExporter(string format)
    {
        return string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase)
            ? new XlsxBomExporter()
            : new CsvBomExporter();
    }

    private static IReadOnlyList<string> CollectHeaders(IEnumerable<BomRow> rows)
    {
        var headers = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var row in rows)
        {
            foreach (var header in row.Values.Keys)
            {
                if (seen.Add(header))
                {
                    headers.Add(header);
                }
            }
        }

        return headers;
    }

    private static string SanitizeFileName(string? fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return "pipe-bom-preview";
        }

        var invalidCharacters = Path.GetInvalidFileNameChars();
        return string.Concat(fileName.Select(character => invalidCharacters.Contains(character) ? '_' : character));
    }

    private static Dictionary<PropertyScope, DataGridView> CreateSelectedPropertyGrids()
    {
        return GetPropertyScopeDisplayOrder()
            .ToDictionary(
                scope => scope,
                _ =>
                {
                    var grid = CreateReadOnlyGrid();
                    grid.DataSource = CreateSelectedPropertiesTable();
                    return grid;
                });
    }

    private static DataTable CreateSelectedPropertiesTable()
    {
        var table = new DataTable();
        table.Columns.Add("Property");
        table.Columns.Add("Evaluated Value");
        table.Columns.Add("Raw Value");
        table.Columns.Add("Source");
        return table;
    }

    private static IReadOnlyList<PropertyScope> GetPropertyScopeDisplayOrder()
    {
        return
        [
            PropertyScope.Configuration,
            PropertyScope.File,
            PropertyScope.Component,
            PropertyScope.CutList,
            PropertyScope.Unknown,
        ];
    }

    private static string GetPropertyScopeTabText(PropertyScope scope)
    {
        return scope switch
        {
            PropertyScope.Configuration => "Configuration",
            PropertyScope.File => "Part File",
            PropertyScope.Component => "Component",
            PropertyScope.CutList => "Cut List",
            _ => "Unknown",
        };
    }

    private sealed class PipeColumnMappingRow
    {
        public string SourceProperty { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public bool Enabled { get; set; } = true;

        public bool GroupBy { get; set; } = true;

        public int Order { get; set; }

        public string Unit { get; set; } = string.Empty;
    }

    private sealed class SectionColumnMappingRow
    {
        public string SourceProperty { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public bool Enabled { get; set; } = true;

        public bool GroupBy { get; set; } = true;

        public int Order { get; set; }

        public string Unit { get; set; } = string.Empty;
    }

    private sealed class SectionRuleMappingRow
    {
        public string SourceProperty { get; set; } = KnownPropertyNames.PrimaryFamily;

        public string MatchValue { get; set; } = string.Empty;

        public string Section { get; set; } = KnownBomSections.Other;
    }

    private sealed class AccessoryMappingRow
    {
        public string SourceProperty { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public string BomSection { get; set; } = KnownBomSections.OtherAccessories;
    }
}
