using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using BomCore;
using SolidWorks.Interop.sldworks;

namespace SolidWorksBOMAddin;

internal sealed class BomPreviewShellForm : Form
{
    private readonly BomPipeAddin _addin;
    private readonly ProfileStore _profileStore = new();
    private readonly PropertyDiscoveryService _propertyDiscoveryService = new();
    private readonly BomGenerator _bomGenerator = new();
    private readonly BindingList<PipeColumnMappingRow> _pipeColumns = [];
    private readonly BindingList<AccessoryMappingRow> _accessoryRules = [];
    private readonly DataGridView _selectedPropertiesGrid;
    private readonly DataGridView _pipeColumnsGrid;
    private readonly DataGridView _accessoryRulesGrid;
    private readonly DataGridView _previewGrid;
    private readonly DataGridView _diagnosticsGrid;
    private readonly ListBox _discoveredPropertiesList;
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
    private int _componentsScanned;
    private int? _componentsSkipped;

    public BomPreviewShellForm(BomPipeAddin addin)
    {
        _addin = addin ?? throw new ArgumentNullException(nameof(addin));

        Text = "Pipe BOM Preview Shell";
        StartPosition = FormStartPosition.CenterScreen;
        Width = 1280;
        Height = 860;
        MinimumSize = new Size(960, 640);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(8),
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
        };

        commandPanel.Controls.Add(CreateActionButton("Read Selected Part Properties", (_, _) => ReadSelectedPartProperties()));
        commandPanel.Controls.Add(CreateActionButton("Scan Active Assembly", (_, _) => ScanActiveAssembly()));
        commandPanel.Controls.Add(CreateActionButton("Edit BOM Mapping", (_, _) => ShowMappingTab()));
        commandPanel.Controls.Add(CreateActionButton("Generate BOM Preview", (_, _) => GenerateBomPreview()));
        commandPanel.Controls.Add(CreateActionButton("Export CSV", (_, _) => ExportBom("csv")));
        commandPanel.Controls.Add(CreateActionButton("Export Excel", (_, _) => ExportBom("xlsx")));

        var infoPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            AutoSize = true,
            Padding = new Padding(0, 4, 0, 4),
        };

        _summaryLabel = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Text = "Scan an active SolidWorks assembly to populate component counts, mapping hints, and preview rows.",
        };
        _profileLabel = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Text = "Mapping note: NumGaskets and NumClamps create accessory rows in Pipe Accessories rather than standard pipe columns.",
        };
        infoPanel.Controls.Add(_summaryLabel, 0, 0);
        infoPanel.Controls.Add(_profileLabel, 0, 1);

        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill,
        };

        _selectedPropertiesGrid = CreateReadOnlyGrid();
        _pipeColumnsGrid = CreatePipeColumnsGrid();
        _accessoryRulesGrid = CreateAccessoryRulesGrid();
        _previewGrid = CreateReadOnlyGrid();
        _diagnosticsGrid = CreateReadOnlyGrid();
        _discoveredPropertiesList = new ListBox
        {
            Dock = DockStyle.Fill,
            IntegralHeight = false,
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
            Text = "Ready.",
        };

        root.Controls.Add(commandPanel, 0, 0);
        root.Controls.Add(infoPanel, 0, 1);
        root.Controls.Add(_tabControl, 0, 2);
        root.Controls.Add(_statusLabel, 0, 3);
        Controls.Add(root);

        LoadDefaultProfileContext();
    }

    private static Button CreateActionButton(string text, EventHandler onClick)
    {
        var button = new Button
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 0, 8, 8),
            Padding = new Padding(10, 6, 10, 6),
            Text = text,
            UseVisualStyleBackColor = true,
        };
        button.Click += onClick;
        return button;
    }

    private static DataGridView CreateReadOnlyGrid()
    {
        return new DataGridView
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

        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.SourceProperty), "Source Property"));
        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.DisplayName), "Display Name"));
        grid.Columns.Add(CreateCheckboxColumn(nameof(PipeColumnMappingRow.Enabled), "Enabled"));
        grid.Columns.Add(CreateCheckboxColumn(nameof(PipeColumnMappingRow.GroupBy), "Group By"));
        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.Order), "Order", 80));
        grid.Columns.Add(CreateTextColumn(nameof(PipeColumnMappingRow.Unit), "Unit", 90));
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

        grid.Columns.Add(CreateTextColumn(nameof(AccessoryMappingRow.SourceProperty), "Accessory Property"));
        grid.Columns.Add(CreateTextColumn(nameof(AccessoryMappingRow.DisplayName), "Accessory Row Description"));
        grid.Columns.Add(CreateTextColumn(nameof(AccessoryMappingRow.BomSection), "BOM Section"));
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
        };
        page.Controls.Add(_selectedPropertiesGrid);
        return page;
    }

    private TabPage CreateMappingTab()
    {
        var page = new TabPage("BOM Mapping")
        {
            Name = "MappingTab",
        };
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 4,
            Padding = new Padding(4),
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240F));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 55F));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));

        var discoveredLabel = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Padding = new Padding(0, 0, 0, 6),
            Text = "Discovered assembly properties",
        };

        var mappingNote = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Padding = new Padding(0, 0, 0, 6),
            Text = "Pipe columns drive grouping. NumGaskets and NumClamps belong in accessory rules so they generate separate accessory rows.",
        };

        var saveMappingButton = new Button
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(10, 6, 10, 6),
            Text = "Save Mapping",
            UseVisualStyleBackColor = true,
        };
        saveMappingButton.Click += (_, _) => SaveMapping();

        var mappingHeader = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
        };
        mappingHeader.Controls.Add(mappingNote);
        mappingHeader.Controls.Add(saveMappingButton);

        var pipeGroup = new GroupBox
        {
            Dock = DockStyle.Fill,
            Text = "Pipe Columns",
            Padding = new Padding(8),
        };
        pipeGroup.Controls.Add(_pipeColumnsGrid);

        var accessoryGroup = new GroupBox
        {
            Dock = DockStyle.Fill,
            Text = "Accessory Rules",
            Padding = new Padding(8),
        };
        accessoryGroup.Controls.Add(_accessoryRulesGrid);

        layout.Controls.Add(discoveredLabel, 0, 0);
        layout.Controls.Add(mappingHeader, 1, 0);
        layout.Controls.Add(_discoveredPropertiesList, 0, 1);
        layout.SetRowSpan(_discoveredPropertiesList, 3);
        layout.Controls.Add(pipeGroup, 1, 1);
        layout.Controls.Add(new Label { AutoSize = true, Text = "Accessory rows emitted from numeric properties" }, 1, 2);
        layout.Controls.Add(accessoryGroup, 1, 3);

        page.Controls.Add(layout);
        return page;
    }

    private TabPage CreatePreviewTab()
    {
        var page = new TabPage("BOM Preview")
        {
            Name = "PreviewTab",
        };
        page.Controls.Add(_previewGrid);
        return page;
    }

    private TabPage CreateDiagnosticsTab()
    {
        var page = new TabPage("Diagnostics")
        {
            Name = "DiagnosticsTab",
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

            var table = new DataTable();
            table.Columns.Add("Name");
            table.Columns.Add("Value");
            table.Columns.Add("Scope");
            table.Columns.Add("Source");

            foreach (var property in properties.OrderBy(property => property.Name, StringComparer.OrdinalIgnoreCase))
            {
                table.Rows.Add(
                    property.Name,
                    property.EffectiveValue,
                    property.Scope.ToString(),
                    property.Source ?? string.Empty);
            }

            _selectedPropertiesGrid.DataSource = table;
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
            _tabControl.SelectedTab = _tabControl.TabPages["PreviewTab"] ?? _tabControl.TabPages[2];
            SetStatus($"Generated {result.Rows.Count} BOM row(s) using {profile.PipeColumns.Count} pipe column rule(s) and {profile.AccessoryRules.Count} accessory rule(s).");
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
        UpdateSummary();
    }

    private ProfileLoadResult LoadEffectiveProfile(string? assemblyPath)
    {
        var options = BuildProfileStoreOptions();
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
            DefaultProfilePath = Path.Combine(AppContext.BaseDirectory, "profiles", "default.pipebom.json"),
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

    private void BindProfile(BomProfile profile, PropertyDiscoveryResult? discovery)
    {
        _pipeColumns.Clear();
        foreach (var column in profile.PipeColumns.OrderBy(column => column.Order))
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

        _accessoryRules.Clear();
        foreach (var accessoryRule in profile.AccessoryRules)
        {
            _accessoryRules.Add(new AccessoryMappingRow
            {
                SourceProperty = accessoryRule.SourceProperty,
                DisplayName = accessoryRule.DisplayName,
                BomSection = string.IsNullOrWhiteSpace(accessoryRule.BomSection) ? KnownBomSections.PipeAccessories : accessoryRule.BomSection,
            });
        }

        _discoveredPropertiesList.BeginUpdate();
        _discoveredPropertiesList.Items.Clear();
        if (discovery is not null)
        {
            foreach (var propertyName in discovery.DiscoveredProperties)
            {
                _discoveredPropertiesList.Items.Add(propertyName);
            }
        }
        _discoveredPropertiesList.EndUpdate();
        UpdateSummary();
    }

    private void BindPreview(BomResult result)
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

        _previewGrid.DataSource = table;
    }

    private void BindDiagnostics(IEnumerable<BomDiagnostic> diagnostics)
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

        _diagnosticsGrid.DataSource = table;
    }

    private void UpdateSummary()
    {
        var assemblyText = string.IsNullOrWhiteSpace(_assemblyDisplayName) ? "No assembly scanned" : _assemblyDisplayName;
        var componentSummary = _componentsScanned > 0
            ? $"Scanned: {_componentsScanned} | Loaded: {_scannedComponents.Count} | Skipped: {_componentsSkipped ?? 0}"
            : $"Loaded: {_scannedComponents.Count}";
        _summaryLabel.Text = $"Assembly: {assemblyText} | {componentSummary}";
        _profileLabel.Text = $"Profile: {_profileSourcePath ?? "Built-in default"} | Accessory reminder: NumGaskets and NumClamps emit accessory rows.";
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
        _accessoryRulesGrid.EndEdit();

        var orderFallback = 1;
        var pipeColumns = _pipeColumns
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

        var accessoryRules = _accessoryRules
            .Where(row => !string.IsNullOrWhiteSpace(row.SourceProperty) && !string.IsNullOrWhiteSpace(row.DisplayName))
            .Select(row => new AccessoryRule
            {
                SourceProperty = row.SourceProperty.Trim(),
                DisplayName = row.DisplayName.Trim(),
                BomSection = string.IsNullOrWhiteSpace(row.BomSection) ? KnownBomSections.PipeAccessories : row.BomSection.Trim(),
            })
            .ToList();

        _currentProfile = _currentProfile with
        {
            PipeColumns = pipeColumns,
            AccessoryRules = accessoryRules,
            ProfileName = string.IsNullOrWhiteSpace(_currentProfile.ProfileName) ? "AFCA Pipe BOM" : _currentProfile.ProfileName,
            Version = _currentProfile.Version <= 0 ? 1 : _currentProfile.Version,
        };

        _profileDiagnostics = BomProfileSerializer.Validate(_currentProfile);
        return _currentProfile;
    }

    private string ResolveProfileSavePath()
    {
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

    private sealed class PipeColumnMappingRow
    {
        public string SourceProperty { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public bool Enabled { get; set; } = true;

        public bool GroupBy { get; set; } = true;

        public int Order { get; set; }

        public string Unit { get; set; } = string.Empty;
    }

    private sealed class AccessoryMappingRow
    {
        public string SourceProperty { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public string BomSection { get; set; } = KnownBomSections.PipeAccessories;
    }
}
