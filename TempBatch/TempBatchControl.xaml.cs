using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace TempBatch
{
    public partial class TempBatchControl : UserControl
    {
        // 取消令牌源，用于取消异步操作
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isGenerating = false;
        private int _generatedFileCount = 0;
        private int _totalFileCount = 0;

        // 类成员
        private DataTable _excelDataTable;
        private List<List<string>> _excelData;
        private List<string> _excelColumnNames;
        private List<KeyValuePair<string, int>> _cachedExcelColumns = new List<KeyValuePair<string, int>>();
        private Dictionary<string, int> _currentFileParamCounters = new Dictionary<string, int>();
        private Dictionary<string, int> _paramBaseStartRows = new Dictionary<string, int>();
        private Dictionary<string, int> _excelParamGlobalCount = new Dictionary<string, int>();
        private Dictionary<string, int> _paramOccurrenceCount = new Dictionary<string, int>();
        private string _templateFileExtension;
        private List<string> _allParameters = new List<string>();
        private List<string> _fileNameParameters = new List<string>();

        // 用于追踪每个Excel列参数的全局位置
        private Dictionary<string, int> _excelParamGlobalPosition = new Dictionary<string, int>();

        // 用于UI操作的Dispatcher
        private Dispatcher _uiDispatcher;

        // 存储UI控件引用（在UI线程初始化）
        private TextBox _filePathTextBox;
        private TextBox _tempFilePathTextBox;
        private ComboBox _sheetComboBox;
        private TextBox _fileCountTextBox;
        private TextBox _fileNameFormatTextBox;
        private CheckBox _mergeFilesCheckBox;
        private StackPanel _parameterMappingPanel;
        private TextBlock _noParametersText;
        private DataGrid _previewDataGrid;
        private TextBlock _statusText;
        private ProgressBar _generationProgressBar;
        private Button _cancelButton;

        public TempBatchControl()
        {
            InitializeComponent();
            // 保存UI线程的Dispatcher
            _uiDispatcher = Dispatcher.CurrentDispatcher;
            // 在UI线程初始化控件引用
            _uiDispatcher.Invoke(InitializeControlReferences);
            InitializeDefaults();
        }

        // 在UI线程初始化所有控件引用
        private void InitializeControlReferences()
        {
            _filePathTextBox = FindName("FilePathTextBox") as TextBox;
            _tempFilePathTextBox = FindName("TempFilePathTextBox") as TextBox;
            _sheetComboBox = FindName("SheetComboBox") as ComboBox;
            _fileCountTextBox = FindName("FileCountTextBox") as TextBox;
            _fileNameFormatTextBox = FindName("FileNameFormatTextBox") as TextBox;
            _mergeFilesCheckBox = FindName("MergeFilesCheckBox") as CheckBox;
            _parameterMappingPanel = FindName("ParameterMappingPanel") as StackPanel;
            _noParametersText = FindName("NoParametersText") as TextBlock;
            _previewDataGrid = FindName("PreviewDataGrid") as DataGrid;
            _statusText = FindName("StatusText") as TextBlock;
            _generationProgressBar = FindName("GenerationProgressBar") as ProgressBar;
            _cancelButton = FindName("CancelButton") as Button;

            // 检查关键控件是否存在
            CheckRequiredControls();
        }

        // 检查必要控件是否存在
        private void CheckRequiredControls()
        {
            var missingControls = new List<string>();
            if (_statusText == null) missingControls.Add("StatusText");
            if (_previewDataGrid == null) missingControls.Add("PreviewDataGrid");
            if (_parameterMappingPanel == null) missingControls.Add("ParameterMappingPanel");
            if (_filePathTextBox == null) missingControls.Add("FilePathTextBox");
            if (_sheetComboBox == null) missingControls.Add("SheetComboBox");

            if (missingControls.Count > 0)
            {
                MessageBox.Show($"缺少必要控件，请检查XAML文件：{string.Join(", ", missingControls)}",
                    "控件缺失", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InitializeDefaults()
        {
            Loaded += (s, e) =>
            {
                if (_fileNameFormatTextBox != null)
                    _fileNameFormatTextBox.Text = "文件_{序号}";

                if (_previewDataGrid != null)
                {
                    _previewDataGrid.AutoGenerateColumns = true;
                    _previewDataGrid.CanUserSortColumns = true;
                    _previewDataGrid.CanUserResizeColumns = true;
                    _previewDataGrid.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
                    _previewDataGrid.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
                    _previewDataGrid.MaxHeight = 500;
                }

                if (_parameterMappingPanel != null && _noParametersText != null && !_parameterMappingPanel.Children.Contains(_noParametersText))
                {
                    _parameterMappingPanel.Children.Add(_noParametersText);
                }

                UpdateStatus("就绪 - 请选择Excel文件和模板文件开始");
            };
        }

        #region 事件处理
        // 选择Excel文件 - 确保在UI线程执行
        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => SelectExcelFile_Click(sender, e));
                return;
            }

            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls",
                Title = "选择Excel文件"
            };

            bool? dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == true && _filePathTextBox != null)
            {
                _filePathTextBox.Text = openFileDialog.FileName;

                // 加载Sheet名称 - 耗时操作放到后台线程
                Task.Run(() => LoadSheetNamesAsync(openFileDialog.FileName));
            }
        }

        // 选择模板文件 - 确保在UI线程执行
        private void SelectTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => SelectTemplateFile_Click(sender, e));
                return;
            }

            var openFileDialog = new OpenFileDialog
            {
                Filter = "所有文件 (*.*)|*.*",
                Title = "选择模板文件"
            };

            bool? dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == true && _tempFilePathTextBox != null)
            {
                _tempFilePathTextBox.Text = openFileDialog.FileName;
                _templateFileExtension = Path.GetExtension(openFileDialog.FileName);
                UpdateStatus($"模板文件加载成功，文件类型: {_templateFileExtension}，请点击参数解析");
            }
        }

        private void TemplateParam_Click(object sender, RoutedEventArgs e)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => TemplateParam_Click(sender, e));
                return;
            }

            if (_tempFilePathTextBox == null || string.IsNullOrEmpty(_tempFilePathTextBox.Text) || !File.Exists(_tempFilePathTextBox.Text))
            {
                UpdateStatus("请选择有效的模板文件");
                return;
            }

            try
            {
                string templateContent = null;
                try
                {
                    templateContent = File.ReadAllText(_tempFilePathTextBox.Text);
                }
                catch (Exception ex)
                {
                    UpdateStatus($"读取模板文件失败：{ex.Message}");
                    return;
                }

                if (string.IsNullOrEmpty(templateContent))
                {
                    UpdateStatus("模板文件内容为空，无法解析参数");
                    return;
                }

                MatchCollection templateMatches = null;
                try
                {
                    var paramRegex = new Regex(@"#\{(.*?)\}");
                    templateMatches = paramRegex.Matches(templateContent);
                }
                catch (ArgumentException ex)
                {
                    UpdateStatus($"正则表达式错误：{ex.Message}");
                    return;
                }

                _fileNameParameters.Clear();
                string fileNameFormat = "文件_{序号}";
                if (_fileNameFormatTextBox != null && !string.IsNullOrEmpty(_fileNameFormatTextBox.Text))
                {
                    fileNameFormat = _fileNameFormatTextBox.Text;
                }

                MatchCollection fileNameMatches = null;
                try
                {
                    var paramRegex = new Regex(@"#\{([a-zA-Z0-9_]+)\}");
                    fileNameMatches = paramRegex.Matches(fileNameFormat);

                    foreach (Match match in fileNameMatches.Cast<Match>())
                    {
                        if (match?.Groups != null && match.Groups.Count > 1)
                        {
                            string paramName = match.Groups[1].Value;
                            if (!string.IsNullOrEmpty(paramName) && !_fileNameParameters.Contains(paramName))
                            {
                                _fileNameParameters.Add(paramName);
                            }
                        }
                    }
                }
                catch (ArgumentException ex)
                {
                    UpdateStatus($"文件名格式正则表达式错误：{ex.Message}");
                    return;
                }

                _allParameters.Clear();
                AddParametersFromMatches(templateMatches);
                AddParametersFromMatches(fileNameMatches);

                // 重新计算参数出现次数
                _paramOccurrenceCount.Clear();
                RecordParameterOccurrences(templateMatches);
                RecordParameterOccurrences(fileNameMatches);

                GenerateParameterControls(_allParameters);
                UpdateStatus($"参数解析成功，共找到 {_allParameters.Count} 个参数（包括文件名中的 {_fileNameParameters.Count} 个参数）");
            }
            catch (Exception ex)
            {
                UpdateStatus($"参数解析失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"参数解析异常详情: {ex.ToString()}");
            }
        }

        private async void GenerateFiles_Click(object sender, RoutedEventArgs e)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => GenerateFiles_Click(sender, e));
                return;
            }

            if (_isGenerating)
            {
                UpdateStatus("正在生成文件，请等待...");
                return;
            }

            var generateButton = sender as Button;
            if (generateButton != null)
                generateButton.IsEnabled = false;

            if (_generationProgressBar != null)
            {
                _generationProgressBar.Visibility = Visibility.Visible;
                _generationProgressBar.Value = 0;
            }

            if (_cancelButton != null)
            {
                _cancelButton.Visibility = Visibility.Visible;
            }

            try
            {
                _isGenerating = true;
                _generatedFileCount = 0;
                _cancellationTokenSource = new CancellationTokenSource();

                // 获取文件数量（在UI线程）
                int fileCount = 0;
                if (_fileCountTextBox != null && !int.TryParse(_fileCountTextBox.Text, out fileCount) || fileCount <= 0)
                {
                    UpdateStatus("请输入有效的生成文件数量");
                    ResetGenerationUI(generateButton);
                    return;
                }
                _totalFileCount = fileCount;

                // 保存当前UI控件的值到局部变量，避免跨线程访问
                string excelFilePath = _filePathTextBox?.Text;
                string templateFilePath = _tempFilePathTextBox?.Text;
                string selectedSheetName = _sheetComboBox?.SelectedItem?.ToString();
                string fileNameFormat = _fileNameFormatTextBox?.Text;
                bool mergeFiles = _mergeFilesCheckBox?.IsChecked ?? false;

                // 重置全局位置追踪器
                _excelParamGlobalPosition.Clear();

                // 使用Task.Run在后台线程执行，传入所需的参数
                await Task.Run(() => GenerateOutputFiles(
                    _cancellationTokenSource.Token,
                    excelFilePath,
                    templateFilePath,
                    selectedSheetName,
                    fileCount,
                    fileNameFormat,
                    mergeFiles
                ), _cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                UpdateStatus("文件生成已取消");
            }
            catch (Exception ex)
            {
                UpdateStatus($"生成文件时发生错误: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"生成文件异常详情: {ex.ToString()}");
            }
            finally
            {
                _isGenerating = false;
                ResetGenerationUI(generateButton);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _cancellationTokenSource.Cancel();
                UpdateStatus("正在取消文件生成...");
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => Clear_Click(sender, e));
                return;
            }

            if (_isGenerating && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
            }

            if (_filePathTextBox != null) _filePathTextBox.Text = string.Empty;
            if (_tempFilePathTextBox != null) _tempFilePathTextBox.Text = string.Empty;
            if (_sheetComboBox != null) _sheetComboBox.Items.Clear();
            if (_fileCountTextBox != null) _fileCountTextBox.Text = string.Empty;
            if (_fileNameFormatTextBox != null) _fileNameFormatTextBox.Text = "文件_{序号}";
            if (_parameterMappingPanel != null)
            {
                _parameterMappingPanel.Children.Clear();
                if (_noParametersText != null)
                    _parameterMappingPanel.Children.Add(_noParametersText);
            }
            if (_mergeFilesCheckBox != null) _mergeFilesCheckBox.IsChecked = false;

            _excelData = null;
            _excelDataTable = null;
            _excelColumnNames = null;
            _allParameters.Clear();
            _paramOccurrenceCount.Clear();
            _currentFileParamCounters.Clear();
            _paramBaseStartRows.Clear();
            _excelParamGlobalCount.Clear();
            _excelParamGlobalPosition.Clear(); // 清除全局位置追踪器
            _fileNameParameters.Clear();

            if (_generationProgressBar != null)
            {
                _generationProgressBar.Visibility = Visibility.Collapsed;
                _generationProgressBar.Value = 0;
            }

            if (_cancelButton != null)
            {
                _cancelButton.Visibility = Visibility.Collapsed;
            }

            UpdateStatus("已清空所有配置");
        }

        // 修复Sheet加载功能
        private void LoadSheet_Click(object sender, RoutedEventArgs e)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => LoadSheet_Click(sender, e));
                return;
            }

            // 验证Excel文件是否选择
            if (_filePathTextBox == null || string.IsNullOrEmpty(_filePathTextBox.Text) || !File.Exists(_filePathTextBox.Text))
            {
                UpdateStatus("请选择有效的Excel文件");
                return;
            }

            // 验证Sheet是否选择
            if (_sheetComboBox == null || _sheetComboBox.SelectedItem == null)
            {
                UpdateStatus("请选择一个Sheet");
                return;
            }

            string sheetName = _sheetComboBox.SelectedItem.ToString();
            string filePath = _filePathTextBox.Text;
            UpdateStatus($"正在加载Sheet: {sheetName}...");

            // 使用带超时的异步加载，避免无限等待
            var cancellationSource = new CancellationTokenSource(30000); // 30秒超时
            Task.Run(() =>
            {
                var data = ReadExcelData(filePath, sheetName, out string errorMsg, out var columnNames);

                _uiDispatcher.Invoke(() =>
                {
                    if (cancellationSource.Token.IsCancellationRequested)
                    {
                        UpdateStatus($"加载Sheet超时（30秒）: {sheetName}");
                        return;
                    }

                    if (data != null && !string.IsNullOrEmpty(errorMsg))
                    {
                        UpdateStatus($"读取Excel数据时警告: {errorMsg}");
                    }
                    else if (data == null)
                    {
                        UpdateStatus($"读取Excel数据失败: {errorMsg}");
                        return;
                    }

                    // 存储加载的数据
                    _excelData = data;
                    _excelColumnNames = columnNames;

                    // 更新数据网格
                    if (_previewDataGrid != null)
                    {
                        _excelDataTable = new DataTable();
                        if (_excelData.Count > 0)
                        {
                            for (int i = 0; i < _excelColumnNames.Count; i++)
                            {
                                string columnName = string.IsNullOrEmpty(_excelColumnNames[i]) ? $"列 {i + 1}" : _excelColumnNames[i];
                                _excelDataTable.Columns.Add(columnName);
                            }

                            // 限制预览行数，防止大数据量卡顿
                            int rowsToAdd = Math.Min(500, _excelData.Count);
                            for (int i = 0; i < rowsToAdd; i++)
                            {
                                DataRow row = _excelDataTable.NewRow();
                                for (int j = 0; j < _excelData[i].Count; j++)
                                {
                                    if (j < row.Table.Columns.Count)
                                        row[j] = _excelData[i][j];
                                }
                                _excelDataTable.Rows.Add(row);
                            }
                        }

                        _previewDataGrid.ItemsSource = _excelDataTable.DefaultView;
                    }

                    // 更新参数映射面板中的Excel列下拉框
                    UpdateExcelColumnCombos();

                    // 显示成功信息
                    UpdateStatus($"成功加载Sheet: {sheetName}，共 {_excelData.Count} 行数据，{_excelColumnNames.Count} 列");
                });
            }, cancellationSource.Token).ContinueWith(task =>
            {
                if (task.IsFaulted)
                {
                    _uiDispatcher.Invoke(() =>
                    {
                        UpdateStatus($"加载Sheet时发生错误: {task.Exception?.InnerException?.Message}");
                        System.Diagnostics.Debug.WriteLine($"加载Sheet异常: {task.Exception?.ToString()}");
                    });
                }
            });
        }

        private void ConfigType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => ConfigType_SelectionChanged(sender, e));
                return;
            }

            var combo = sender as ComboBox;
            if (combo?.Tag is Tuple<CheckBox, TextBox, ComboBox, ComboBox, TextBox> controls)
            {
                var (paramEnableCheck, valueTextBox, columnCombo, excelModeCombo, emptyValueTextBox) = controls;

                if (paramEnableCheck == null || valueTextBox == null || columnCombo == null || excelModeCombo == null || emptyValueTextBox == null)
                    return;

                if (paramEnableCheck.IsChecked == false)
                {
                    valueTextBox.IsEnabled = false;
                    columnCombo.IsEnabled = false;
                    excelModeCombo.IsEnabled = false;
                    emptyValueTextBox.IsEnabled = false;
                    return;
                }

                string selectedType = (combo.SelectedItem as ComboBoxItem)?.Content.ToString();

                switch (selectedType)
                {
                    case "定值":
                        valueTextBox.IsEnabled = true;
                        columnCombo.IsEnabled = false;
                        excelModeCombo.IsEnabled = false;
                        emptyValueTextBox.IsEnabled = false;
                        break;

                    case "变值":
                        valueTextBox.IsEnabled = true;
                        columnCombo.IsEnabled = false;
                        excelModeCombo.IsEnabled = false;
                        emptyValueTextBox.IsEnabled = false;
                        break;

                    case "Excel列":
                        valueTextBox.IsEnabled = false;
                        columnCombo.IsEnabled = true;
                        excelModeCombo.IsEnabled = true;
                        emptyValueTextBox.IsEnabled = true;
                        // 为空值替换添加默认值，防止空值判断错误
                        if (string.IsNullOrEmpty(emptyValueTextBox.Text))
                        {
                            emptyValueTextBox.Text = ""; // 明确设置为空字符串而非null
                        }
                        break;
                }
            }
        }

        private void IntegerTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private void IntegerTextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string));
                if (!IsTextAllowed(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }
        #endregion

        #region 辅助方法
        // 异步加载Sheet名称，避免阻塞UI
        private void LoadSheetNamesAsync(string filePath)
        {
            try
            {
                var sheetNames = new List<string>();
                string errorMsg = string.Empty;

                // 在后台线程读取Sheet名称
                using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = spreadsheet.WorkbookPart;
                    if (workbookPart == null)
                    {
                        errorMsg = "Excel文件中未找到工作簿";
                        _uiDispatcher.Invoke(() => UpdateStatus(errorMsg));
                        return;
                    }

                    var sheets = workbookPart.Workbook.Sheets.Elements<Sheet>();
                    foreach (var sheet in sheets)
                    {
                        sheetNames.Add(sheet.Name);
                    }
                }

                // 在UI线程更新SheetComboBox
                _uiDispatcher.Invoke(() =>
                {
                    if (_sheetComboBox == null)
                    {
                        UpdateStatus("SheetComboBox控件不存在，无法加载Sheet名称");
                        return;
                    }

                    _sheetComboBox.Items.Clear();
                    foreach (var sheetName in sheetNames)
                    {
                        _sheetComboBox.Items.Add(sheetName);
                    }

                    if (sheetNames.Count > 0)
                    {
                        _sheetComboBox.SelectedIndex = 0;
                        UpdateStatus($"成功加载 {sheetNames.Count} 个Sheet");
                    }
                    else
                    {
                        UpdateStatus("Excel文件中未找到任何Sheet");
                    }
                });
            }
            catch (IOException ex)
            {
                _uiDispatcher.Invoke(() =>
                    UpdateStatus($"文件IO错误: 可能文件已被打开 - {ex.Message}")
                );
            }
            catch (Exception ex)
            {
                _uiDispatcher.Invoke(() =>
                    UpdateStatus($"加载Sheet名称失败: {ex.Message}")
                );
                System.Diagnostics.Debug.WriteLine($"加载Sheet名称异常: {ex.ToString()}");
            }
        }

        private void AddParametersFromMatches(MatchCollection matches)
        {
            if (matches == null) return;

            foreach (Match match in matches.Cast<Match>())
            {
                if (match?.Groups != null && match.Groups.Count > 1)
                {
                    string paramName = match.Groups[1].Value;
                    if (!string.IsNullOrEmpty(paramName) && !_allParameters.Contains(paramName))
                    {
                        _allParameters.Add(paramName);
                    }
                }
            }
        }

        // 精确记录参数出现次数
        private void RecordParameterOccurrences(MatchCollection matches)
        {
            if (matches == null) return;

            foreach (Match match in matches.Cast<Match>())
            {
                if (match?.Groups != null && match.Groups.Count > 1)
                {
                    string paramName = match.Groups[1].Value;
                    if (!string.IsNullOrEmpty(paramName))
                    {
                        if (_paramOccurrenceCount.ContainsKey(paramName))
                            _paramOccurrenceCount[paramName]++;
                        else
                            _paramOccurrenceCount[paramName] = 1;
                    }
                }
            }
        }

        private void GenerateParameterControls(List<string> parameters)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => GenerateParameterControls(parameters));
                return;
            }

            if (_parameterMappingPanel == null)
            {
                UpdateStatus("ParameterMappingPanel控件不存在，无法生成参数控件");
                return;
            }

            _parameterMappingPanel.Children.Clear();

            if (parameters == null || parameters.Count == 0)
            {
                if (_noParametersText != null)
                    _parameterMappingPanel.Children.Add(_noParametersText);
                return;
            }

            foreach (var param in parameters)
            {
                var paramPanel = new StackPanel { Margin = new Thickness(0, 5, 0, 5) };
                var grid = new Grid();

                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });

                var paramEnableCheck = new CheckBox
                {
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center,
                    IsChecked = true
                };
                Grid.SetColumn(paramEnableCheck, 0);

                var paramNameText = new TextBlock
                {
                    Text = param,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5),
                    TextWrapping = TextWrapping.NoWrap
                };
                // 显示参数出现次数
                if (_paramOccurrenceCount.TryGetValue(param, out int count))
                {
                    paramNameText.ToolTip = $"在模板中出现 {count} 次";
                }
                Grid.SetColumn(paramNameText, 1);

                var configTypeCombo = new ComboBox
                {
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center
                };
                configTypeCombo.Items.Add(new ComboBoxItem { Content = "定值" });
                configTypeCombo.Items.Add(new ComboBoxItem { Content = "变值" });
                configTypeCombo.Items.Add(new ComboBoxItem { Content = "Excel列" });
                configTypeCombo.SelectedIndex = 0;
                configTypeCombo.SelectionChanged += ConfigType_SelectionChanged;
                Grid.SetColumn(configTypeCombo, 2);

                var configValueTextBox = new TextBox
                {
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(configValueTextBox, 3);

                var excelColumnCombo = new ComboBox
                {
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center,
                    IsEnabled = false
                };
                Grid.SetColumn(excelColumnCombo, 4);

                var excelModeCombo = new ComboBox
                {
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center,
                    IsEnabled = false
                };
                excelModeCombo.Items.Add(new ComboBoxItem { Content = "整个文件使用同一行" });
                excelModeCombo.Items.Add(new ComboBoxItem { Content = "顺序填写" });
                excelModeCombo.SelectedIndex = 0;
                Grid.SetColumn(excelModeCombo, 5);

                var emptyValueTextBox = new TextBox
                {
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center,
                    IsEnabled = false,
                    ToolTip = "遇到空数据时使用此值填充",
                    Text = "" // 设置默认空字符串，避免null
                };
                try
                {
                    var emptyValueAdorner = new WatermarkAdorner(emptyValueTextBox, "空值替换内容");
                    AdornerLayer.GetAdornerLayer(emptyValueTextBox)?.Add(emptyValueAdorner);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"为控件添加水印失败: {ex.Message}");
                }
                Grid.SetColumn(emptyValueTextBox, 6);

                configTypeCombo.Tag = new Tuple<CheckBox, TextBox, ComboBox, ComboBox, TextBox>(
                    paramEnableCheck,
                    configValueTextBox,
                    excelColumnCombo,
                    excelModeCombo,
                    emptyValueTextBox
                );

                grid.Children.Add(paramEnableCheck);
                grid.Children.Add(paramNameText);
                grid.Children.Add(configTypeCombo);
                grid.Children.Add(configValueTextBox);
                grid.Children.Add(excelColumnCombo);
                grid.Children.Add(excelModeCombo);
                grid.Children.Add(emptyValueTextBox);

                paramPanel.Children.Add(grid);
                _parameterMappingPanel.Children.Add(paramPanel);
            }

            UpdateExcelColumnCombos();
        }

        private void UpdateExcelColumnCombos()
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => UpdateExcelColumnCombos());
                return;
            }

            if (_excelData == null || _excelData.Count == 0 || _excelColumnNames == null)
                return;

            _cachedExcelColumns.Clear();
            for (int i = 0; i < _excelColumnNames.Count; i++)
            {
                string columnName = string.IsNullOrEmpty(_excelColumnNames[i]) ? $"列 {i + 1}" : _excelColumnNames[i];
                _cachedExcelColumns.Add(new KeyValuePair<string, int>(columnName, i));
            }

            if (_parameterMappingPanel == null) return;

            foreach (var child in _parameterMappingPanel.Children)
            {
                if (child is StackPanel paramPanel && paramPanel.Children[0] is Grid grid)
                {
                    if (grid.Children[4] is ComboBox columnCombo)
                    {
                        columnCombo.Items.Clear();
                        foreach (var column in _cachedExcelColumns)
                        {
                            columnCombo.Items.Add(column);
                        }
                        columnCombo.DisplayMemberPath = "Key";
                        columnCombo.SelectedValuePath = "Value";
                    }
                }
            }
        }

        // 辅助类用于从ParseParamConfigs返回结果和错误信息
        private class ParamConfigResult
        {
            public List<ParamConfig> Configs { get; set; }
            public string ErrorMessage { get; set; }
        }

        // 修改ParseParamConfigs方法，不使用out参数，而是返回包含错误信息的对象
        private ParamConfigResult ParseParamConfigs()
        {
            var result = new ParamConfigResult();
            result.Configs = new List<ParamConfig>();
            result.ErrorMessage = string.Empty;

            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                return _uiDispatcher.Invoke(() => ParseParamConfigs());
            }

            if (_parameterMappingPanel == null)
            {
                result.ErrorMessage = "ParameterMappingPanel控件不存在，无法解析参数配置";
                result.Configs = null;
                return result;
            }

            foreach (var child in _parameterMappingPanel.Children)
            {
                if (child is StackPanel paramPanel && paramPanel.Children[0] is Grid grid)
                {
                    var config = new ParamConfig();

                    if (grid.Children[0] is CheckBox enableCheck)
                        config.IsEnable = enableCheck.IsChecked ?? false;
                    else
                    {
                        result.ErrorMessage = "参数配置控件格式错误";
                        result.Configs = null;
                        return result;
                    }

                    if (grid.Children[1] is TextBlock nameText)
                        config.ParamName = nameText.Text;
                    else
                    {
                        result.ErrorMessage = "参数名称控件格式错误";
                        result.Configs = null;
                        return result;
                    }

                    if (grid.Children[2] is ComboBox typeCombo && typeCombo.SelectedItem is ComboBoxItem typeItem)
                        config.ConfigType = typeItem.Content.ToString();
                    else
                    {
                        result.ErrorMessage = $"参数【{config.ParamName}】的配置类型控件格式错误";
                        result.Configs = null;
                        return result;
                    }

                    if (grid.Children[3] is TextBox valueBox)
                        config.FixedValue = valueBox.Text;
                    else
                    {
                        result.ErrorMessage = $"参数【{config.ParamName}】的配置值控件格式错误";
                        result.Configs = null;
                        return result;
                    }

                    if (grid.Children[4] is ComboBox columnCombo && columnCombo.SelectedItem is KeyValuePair<string, int> column)
                        config.ExcelColumnIndex = column.Value;
                    else if (config.ConfigType == "Excel列")
                    {
                        result.ErrorMessage = $"参数【{config.ParamName}】未选择Excel列";
                        result.Configs = null;
                        return result;
                    }

                    if (grid.Children[5] is ComboBox modeCombo && modeCombo.SelectedItem is ComboBoxItem modeItem)
                        config.ExcelMode = modeItem.Content.ToString();
                    else if (config.ConfigType == "Excel列")
                    {
                        result.ErrorMessage = $"参数【{config.ParamName}】的Excel模式控件格式错误";
                        result.Configs = null;
                        return result;
                    }

                    if (grid.Children[6] is TextBox emptyValueBox)
                        config.EmptyValue = emptyValueBox.Text ?? ""; // 确保不是null
                    else if (config.ConfigType == "Excel列")
                    {
                        result.ErrorMessage = $"参数【{config.ParamName}】的空值填充控件格式错误";
                        result.Configs = null;
                        return result;
                    }

                    // 验证参数配置
                    if (config.IsEnable && !config.IsValid(out string msg))
                    {
                        result.ErrorMessage = msg;
                        result.Configs = null;
                        return result;
                    }

                    if (config.IsEnable)
                        result.Configs.Add(config);
                }
            }

            return result;
        }

        private int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;
            var columnPart = new string(cellReference.Where(char.IsLetter).ToArray());
            int index = 0;
            foreach (char ch in columnPart)
            {
                index = index * 26 + (ch - 'A' + 1);
            }
            return index - 1;
        }

        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty; // 返回空字符串而非null

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (int.TryParse(value, out int index) && sharedStringTable != null && index < sharedStringTable.Count)
                {
                    value = sharedStringTable.ElementAt(index).InnerText;
                }
            }

            return value ?? string.Empty; // 确保不会返回null
        }

        private Encoding GetFileEncoding(string filePath)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = new StreamReader(stream, true))
            {
                reader.ReadToEnd();
                return reader.CurrentEncoding;
            }
        }

        private string ReplaceFirstOccurrence(string source, string find, string replace)
        {
            int index = source.IndexOf(find);
            if (index == -1)
                return source;
            return source.Substring(0, index) + replace + source.Substring(index + find.Length);
        }

        private void UpdateStatus(string message)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => UpdateStatus(message));
                return;
            }

            if (_statusText == null)
            {
                System.Diagnostics.Debug.WriteLine($"状态更新: {message}");
                return;
            }

            _statusText.Text = $"[{DateTime.Now:HH:mm:ss}] {message}";
            // 滚动到最新状态
            _statusText.BringIntoView();
        }

        private void UpdateProgress(int value)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => UpdateProgress(value));
                return;
            }

            if (_generationProgressBar == null) return;

            _generationProgressBar.Value = value;
        }

        private void ShowError(string message)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => ShowError(message));
                return;
            }

            MessageBox.Show(message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Diagnostics.Debug.WriteLine($"错误: {message}");
        }

        private static bool IsTextAllowed(string text)
        {
            return Regex.IsMatch(text, @"^\d*$");
        }

        private string GenerateFileName(string format, int fileIndex, List<ParamConfig> paramConfigs, List<List<string>> excelData)
        {
            string fileName = format ?? "文件_{序号}";

            fileName = fileName.Replace("{序号}", fileIndex.ToString());

            var paramRegex = new Regex(@"#\{([a-zA-Z0-9_]+)\}");
            var matches = paramRegex.Matches(fileName);

            foreach (Match match in matches.Cast<Match>().OrderByDescending(m => m.Length))
            {
                if (match?.Groups == null || match.Groups.Count <= 1)
                    continue;

                string paramName = match.Groups[1].Value;
                var config = paramConfigs.FirstOrDefault(p => p.ParamName == paramName && p.IsEnable);

                if (config != null)
                {
                    bool skipExcelRow;
                    // 检查字典中是否存在该参数
                    if (!_paramBaseStartRows.TryGetValue(paramName, out int paramBaseRow))
                    {
                        UpdateStatus($"警告：参数 {paramName} 不在参数起始行字典中，使用默认值0");
                        paramBaseRow = 0;
                    }

                    string paramValue = GenerateFileNameParamValue(config, excelData, fileIndex - 1, paramBaseRow, out skipExcelRow);

                    if (skipExcelRow)
                    {
                        paramValue = config.EmptyValue;
                    }

                    fileName = fileName.Replace(match.Value, paramValue ?? string.Empty);
                }
                else if (_fileNameParameters.Contains(paramName))
                {
                    UpdateStatus($"警告：文件名中的参数 #{paramName} 未配置映射");
                    fileName = fileName.Replace(match.Value, $"[未配置:{paramName}]");
                }
            }

            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(invalidChar, '_');
            }

            if (!string.IsNullOrEmpty(_templateFileExtension) && !fileName.EndsWith(_templateFileExtension))
            {
                fileName += _templateFileExtension;
            }

            return fileName;
        }

        private string GenerateFileNameParamValue(ParamConfig config, List<List<string>> excelData,
                                                int fileIndex, int paramBaseRow, out bool skipExcelRow)
        {
            skipExcelRow = false;
            string value = string.Empty;

            switch (config.ConfigType)
            {
                case "Excel列":
                    int dataRowIndex = paramBaseRow;

                    // 对于顺序填写模式，使用全局位置追踪器
                    if (config.ExcelMode == "顺序填写")
                    {
                        // 确保全局位置追踪器中存在该参数
                        int globalPos = 0;
                        if (_excelParamGlobalPosition.TryGetValue(config.ParamName, out int existingPos))
                        {
                            globalPos = existingPos;
                        }
                        else
                        {
                            _excelParamGlobalPosition[config.ParamName] = 0;
                        }
                        dataRowIndex = globalPos;
                    }

                    // 验证数据行索引有效性
                    if (dataRowIndex < 0 || dataRowIndex >= excelData.Count)
                    {
                        skipExcelRow = true;
                        UpdateStatus($"警告：参数 {config.ParamName} 的行索引 {dataRowIndex} 超出范围，使用空值替换");
                        return config.EmptyValue;
                    }

                    // 验证列索引有效性
                    if (config.ExcelColumnIndex < 0 || config.ExcelColumnIndex >= excelData[dataRowIndex].Count)
                    {
                        skipExcelRow = true;
                        UpdateStatus($"警告：参数 {config.ParamName} 的列索引 {config.ExcelColumnIndex} 超出范围，使用空值替换");
                        return config.EmptyValue;
                    }

                    value = excelData[dataRowIndex][config.ExcelColumnIndex];

                    // 更精确的空值判断：只在真正为空或空白时才使用替换值
                    if (string.IsNullOrWhiteSpace(value) && !string.IsNullOrEmpty(config.EmptyValue))
                    {
                        skipExcelRow = true;
                        UpdateStatus($"参数 {config.ParamName} 在行 {dataRowIndex + 1} 为空，使用替换值");
                    }
                    break;

                case "定值":
                    value = config.FixedValue ?? string.Empty;
                    break;

                case "变值":
                    if (config.FixedValue.Contains("+"))
                    {
                        var parts = config.FixedValue.Split('+');
                        if (int.TryParse(parts[0], out int start) && int.TryParse(parts[1], out int step))
                        {
                            value = (start + fileIndex * step).ToString();
                        }
                        else
                        {
                            value = "格式错误";
                            UpdateStatus($"警告：参数 {config.ParamName} 的变值格式错误");
                        }
                    }
                    else
                    {
                        value = "格式错误";
                        UpdateStatus($"警告：参数 {config.ParamName} 的变值格式错误");
                    }
                    break;
            }

            return value ?? string.Empty;
        }

        private string GetUniqueFilePath(string filePath)
        {
            if (!File.Exists(filePath))
                return filePath;

            string directory = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            int counter = 1;
            string newFilePath;

            do
            {
                newFilePath = Path.Combine(directory, $"{fileName}_{counter}{extension}");
                counter++;
            } while (File.Exists(newFilePath));

            return newFilePath;
        }

        private void ResetGenerationUI(Button generateButton)
        {
            // 确保在UI线程执行
            if (!_uiDispatcher.CheckAccess())
            {
                _uiDispatcher.Invoke(() => ResetGenerationUI(generateButton));
                return;
            }

            if (generateButton != null)
                generateButton.IsEnabled = true;

            if (_generationProgressBar != null)
            {
                _generationProgressBar.Visibility = Visibility.Collapsed;
            }

            if (_cancelButton != null)
            {
                _cancelButton.Visibility = Visibility.Collapsed;
            }
        }
        #endregion

        #region 核心逻辑
        // 生成文件的核心方法，所有UI值通过参数传入，避免跨线程访问UI控件
        private void GenerateOutputFiles(CancellationToken cancellationToken,
                                        string excelFilePath,
                                        string templateFilePath,
                                        string selectedSheetName,
                                        int fileCount,
                                        string fileNameFormat,
                                        bool mergeFiles)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                // 验证输入参数
                if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
                {
                    UpdateStatus("请选择有效的Excel文件");
                    return;
                }

                if (string.IsNullOrEmpty(selectedSheetName))
                {
                    UpdateStatus("请选择Excel Sheet并点击加载");
                    return;
                }

                if (string.IsNullOrEmpty(templateFilePath) || !File.Exists(templateFilePath))
                {
                    UpdateStatus("请选择有效的模板文件");
                    return;
                }

                if (fileCount <= 0)
                {
                    UpdateStatus("请输入有效的生成文件数量");
                    return;
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 检查文件名参数配置
                if (_fileNameParameters.Count > 0)
                {
                    string errorParams = string.Empty;
                    var configResult = ParseParamConfigs();
                    var configs = configResult.Configs;

                    foreach (var param in _fileNameParameters)
                    {
                        if (configs == null || !configs.Any(c => c.ParamName == param && c.IsEnable))
                        {
                            errorParams += $"{param}, ";
                        }
                    }

                    if (!string.IsNullOrEmpty(errorParams))
                    {
                        UpdateStatus($"警告：文件名中的以下参数未配置或未启用：{errorParams.TrimEnd(',', ' ')}");

                        bool? continueGenerating = false;
                        _uiDispatcher.Invoke(() =>
                        {
                            continueGenerating = (MessageBox.Show(
                                $"文件名中的以下参数未配置或未启用：{errorParams.TrimEnd(',', ' ')}\n是否继续生成文件？",
                                "参数配置不完整",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning) == MessageBoxResult.Yes);
                        });

                        if (continueGenerating != true)
                        {
                            return;
                        }
                    }
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 选择输出路径（在UI线程执行）
                string outputDir = string.Empty;
                _uiDispatcher.Invoke(() =>
                {
                    var dialog = new OpenFileDialog
                    {
                        ValidateNames = false,
                        CheckFileExists = false,
                        CheckPathExists = true,
                        FileName = "选择文件夹",
                        Title = "请选择输出文件夹"
                    };

                    if (!string.IsNullOrEmpty(templateFilePath) && Directory.Exists(Path.GetDirectoryName(templateFilePath)))
                    {
                        dialog.InitialDirectory = Path.GetDirectoryName(templateFilePath);
                    }
                    else if (!string.IsNullOrEmpty(excelFilePath) && Directory.Exists(Path.GetDirectoryName(excelFilePath)))
                    {
                        dialog.InitialDirectory = Path.GetDirectoryName(excelFilePath);
                    }
                    else
                    {
                        dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }

                    bool? dialogResult = dialog.ShowDialog();
                    if (dialogResult == true)
                    {
                        outputDir = Path.GetDirectoryName(dialog.FileName);
                    }
                });

                if (string.IsNullOrEmpty(outputDir) || !Directory.Exists(outputDir))
                {
                    UpdateStatus("选择的文件夹路径无效");
                    return;
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 解析参数配置（使用新的返回对象，避免out参数）
                var paramConfigResult = ParseParamConfigs();
                if (paramConfigResult.Configs == null || !string.IsNullOrEmpty(paramConfigResult.ErrorMessage))
                {
                    UpdateStatus($"参数配置解析失败：{paramConfigResult.ErrorMessage}");
                    return;
                }

                var paramConfigs = paramConfigResult.Configs;
                if (paramConfigs.Count == 0)
                {
                    UpdateStatus("无启用的参数，请至少勾选一个参数");
                    return;
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 确保Excel数据已加载
                if (_excelData == null)
                {
                    _excelData = ReadExcelData(excelFilePath, selectedSheetName, out string errorMsg, out _excelColumnNames);
                    if (_excelData == null)
                    {
                        UpdateStatus($"读取Excel数据失败：{errorMsg}");
                        return;
                    }
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 初始化顺序填写参数的全局位置追踪器
                foreach (var config in paramConfigs.Where(c => c.IsEnable && c.ConfigType == "Excel列" && c.ExcelMode == "顺序填写"))
                {
                    if (!_excelParamGlobalPosition.ContainsKey(config.ParamName))
                    {
                        _excelParamGlobalPosition[config.ParamName] = 0;
                        UpdateStatus($"初始化参数 {config.ParamName} 顺序填写起始位置：0");
                    }
                }

                // 计算参数偏移量和所需行数
                CalculateParamOffsets(paramConfigs, fileCount);

                // 检查数据是否足够
                if (!CheckAllParamsDataEnough(paramConfigs, fileCount))
                {
                    return;
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 读取模板文件内容和编码（每次生成文件都使用原始模板，避免累积替换）
                string originalTemplateContent = File.ReadAllText(templateFilePath);
                var templateEncoding = GetFileEncoding(templateFilePath);
                int successCount = 0;
                var mergedContent = new StringBuilder();

                // 初始化全局计数器
                _excelParamGlobalCount.Clear();
                foreach (var config in paramConfigs.Where(c => c.IsEnable && c.ConfigType == "Excel列"))
                {
                    // 确保键存在
                    if (!_excelParamGlobalCount.ContainsKey(config.ParamName))
                        _excelParamGlobalCount[config.ParamName] = 0;
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 生成每个文件
                for (int fileIndex = 0; fileIndex < fileCount; fileIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 重置当前文件的参数计数器
                    _currentFileParamCounters.Clear();
                    foreach (var config in paramConfigs.Where(c => c.IsEnable && c.ConfigType == "Excel列"))
                    {
                        // 确保键存在
                        if (!_currentFileParamCounters.ContainsKey(config.ParamName))
                            _currentFileParamCounters[config.ParamName] = 0;
                    }

                    // 更新参数起始行
                    UpdateParamBaseStartRows(paramConfigs, fileIndex);

                    // 每次都使用原始模板内容，避免之前的替换影响
                    string currentContent = originalTemplateContent;
                    bool hasBlank = false;
                    bool allParamsEmpty = true; // 跟踪是否所有参数都为空

                    // 处理每个参数
                    foreach (var config in paramConfigs)
                    {
                        if (!config.IsEnable) continue;

                        // 检查参数出现次数字典中是否存在该参数
                        if (!_paramOccurrenceCount.TryGetValue(config.ParamName, out int paramOccurrence))
                        {
                            paramOccurrence = 1; // 使用默认值
                            UpdateStatus($"警告：参数 {config.ParamName} 不在出现次数字典中，使用默认值1");
                        }

                        for (int occurrenceIndex = 0; occurrenceIndex < paramOccurrence; occurrenceIndex++)
                        {
                            bool skipExcelRow;
                            // 检查参数起始行字典中是否存在该参数
                            if (!_paramBaseStartRows.TryGetValue(config.ParamName, out int paramBaseRow))
                            {
                                paramBaseRow = 0; // 使用默认值
                                UpdateStatus($"警告：参数 {config.ParamName} 不在起始行字典中，使用默认值0");
                            }

                            string paramValue = GenerateParamValue(
                                config, _excelData, fileIndex, paramBaseRow, occurrenceIndex,
                                out skipExcelRow
                            );

                            // 如果参数值非空，则不是所有参数都为空
                            if (!string.IsNullOrEmpty(paramValue))
                            {
                                allParamsEmpty = false;
                            }

                            if (skipExcelRow)
                            {
                                hasBlank = true;
                                paramValue = config.EmptyValue;
                            }
                            else if (string.IsNullOrEmpty(paramValue))
                            {
                                paramValue = config.EmptyValue;
                            }

                            // 替换参数占位符
                            string placeholder1 = $"#{config.ParamName}";
                            string placeholder2 = $"#{{{config.ParamName}}}";

                            // 先替换完整格式，再替换简化格式
                            if (currentContent.Contains(placeholder2))
                            {
                                currentContent = ReplaceFirstOccurrence(currentContent, placeholder2, paramValue);
                            }
                            else if (currentContent.Contains(placeholder1))
                            {
                                currentContent = ReplaceFirstOccurrence(currentContent, placeholder1, paramValue);
                            }
                        }
                    }

                    // 生成文件名
                    string fileName = GenerateFileName(fileNameFormat, fileIndex + 1, paramConfigs, _excelData);
                    string outputPath = Path.Combine(outputDir, fileName);
                    outputPath = GetUniqueFilePath(outputPath);

                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        // 即使所有参数都为空，也生成文件（包含原始模板内容）
                        File.WriteAllText(outputPath, currentContent, templateEncoding);
                        successCount++;

                        if (mergeFiles)
                        {
                            mergedContent.AppendLine($"===== {fileName} =====");
                            mergedContent.AppendLine(currentContent);
                            mergedContent.AppendLine();
                        }

                        // 更新进度
                        _generatedFileCount++;
                        int progress = (int)((double)_generatedFileCount / _totalFileCount * 100);
                        UpdateProgress(progress);

                        // 状态信息
                        if (allParamsEmpty)
                        {
                            UpdateStatus(string.Format("生成文件 {0} ({1}/{2}) - 注意：所有参数值为空",
                                fileName, _generatedFileCount, _totalFileCount));
                        }
                        else
                        {
                            UpdateStatus(string.Format("生成文件 {0} ({1}/{2}){3}",
                                fileName, _generatedFileCount, _totalFileCount,
                                hasBlank ? "（注意：部分参数使用了空值替换）" : ""));
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"生成文件 {fileName} 失败: {ex.Message}");
                    }

                    // 短暂延迟，让UI有机会更新
                    Thread.Sleep(10);
                }

                cancellationToken.ThrowIfCancellationRequested();

                // 生成合并文件
                if (mergeFiles && successCount > 0)
                {
                    string mergeFileName = $"ALL_INF{_templateFileExtension}";
                    string mergeFilePath = Path.Combine(outputDir, mergeFileName);
                    mergeFilePath = GetUniqueFilePath(mergeFilePath);

                    try
                    {
                        File.WriteAllText(mergeFilePath, mergedContent.ToString(), templateEncoding);
                        UpdateStatus($"生成合并文件: {mergeFileName}");
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"生成合并文件失败: {ex.Message}");
                    }
                }

                UpdateStatus($"生成完成！共生成 {successCount} 个文件，输出路径：{outputDir}");

                // 打开输出目录
                _uiDispatcher.Invoke(() =>
                {
                    System.Diagnostics.Process.Start("explorer.exe", outputDir);
                });
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                UpdateStatus($"生成文件时发生错误: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"生成文件异常详情: {ex.ToString()}");
            }
        }

        private void CalculateParamOffsets(List<ParamConfig> paramConfigs, int fileCount)
        {
            foreach (var config in paramConfigs.Where(c => c.IsEnable && c.ConfigType == "Excel列"))
            {
                // 确保参数出现次数字典中存在该参数
                if (!_paramOccurrenceCount.TryGetValue(config.ParamName, out int paramOccurrence))
                {
                    paramOccurrence = 1; // 添加默认值
                    _paramOccurrenceCount[config.ParamName] = paramOccurrence;
                    UpdateStatus($"警告：参数 {config.ParamName} 不在出现次数字典中，添加默认值1");
                }

                // 计算每行文件需要消耗的行数
                int rowsPerFile = config.ExcelMode == "整个文件使用同一行" ? 1 : paramOccurrence;

                // 计算总需要的行数
                config.TotalRequiredRows = rowsPerFile * fileCount;
                config.RowOffsetPerFile = rowsPerFile;

                // 对于顺序填写模式，验证Excel数据是否足够
                if (config.ExcelMode == "顺序填写" && _excelData != null && config.TotalRequiredRows > _excelData.Count)
                {
                    UpdateStatus($"警告：参数 {config.ParamName}（顺序填写）需要 {config.TotalRequiredRows} 行数据，但Excel中只有 {_excelData.Count} 行");
                }
            }
        }

        private bool CheckAllParamsDataEnough(List<ParamConfig> paramConfigs, int fileCount)
        {
            foreach (var config in paramConfigs.Where(c => c.IsEnable && c.ConfigType == "Excel列"))
            {
                if (_excelData == null || config.TotalRequiredRows > _excelData.Count)
                {
                    UpdateStatus($"参数【{config.ParamName}】数据不足，需要 {config.TotalRequiredRows} 行，但只找到 {(_excelData?.Count ?? 0)} 行");

                    // 询问用户是否继续
                    bool? continueGenerating = false;
                    _uiDispatcher.Invoke(() =>
                    {
                        continueGenerating = (MessageBox.Show(
                            $"参数【{config.ParamName}】数据不足，需要 {config.TotalRequiredRows} 行，但只找到 {(_excelData?.Count ?? 0)} 行\n是否继续生成文件？",
                            "数据不足",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Warning) == MessageBoxResult.Yes);
                    });

                    if (continueGenerating != true)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        private void UpdateParamBaseStartRows(List<ParamConfig> paramConfigs, int fileIndex)
        {
            _paramBaseStartRows.Clear();
            foreach (var config in paramConfigs.Where(c => c.IsEnable && c.ConfigType == "Excel列"))
            {
                // 对于"整个文件使用同一行"模式，计算每行文件的起始行
                if (config.ExcelMode == "整个文件使用同一行")
                {
                    _paramBaseStartRows[config.ParamName] = fileIndex * config.RowOffsetPerFile;
                    UpdateStatus($"参数 {config.ParamName}（整行模式）第 {fileIndex + 1} 个文件使用行索引: {_paramBaseStartRows[config.ParamName]}");
                }
                // 对于"顺序填写"模式，使用全局位置追踪器，不在这里设置
            }
        }

        private List<List<string>> ReadExcelData(string filePath, string sheetName, out string errorMsg, out List<string> columnNames)
        {
            errorMsg = string.Empty;
            columnNames = new List<string>();
            var result = new List<List<string>>();

            try
            {
                if (!File.Exists(filePath))
                {
                    errorMsg = "Excel文件不存在";
                    return null;
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = spreadsheet.WorkbookPart;
                    if (workbookPart == null)
                    {
                        errorMsg = "无法打开Excel工作簿";
                        return null;
                    }

                    var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(s => s.Name == sheetName);
                    if (sheet == null)
                    {
                        errorMsg = $"未找到Sheet: {sheetName}";
                        return null;
                    }

                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                    if (worksheetPart == null)
                    {
                        errorMsg = $"无法加载Sheet: {sheetName}";
                        return null;
                    }

                    var rows = worksheetPart.Worksheet.Descendants<Row>().ToList();
                    if (rows.Count == 0)
                    {
                        errorMsg = "Sheet中没有数据行";
                        return result;
                    }

                    var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

                    // 读取表头行
                    var headerRow = rows[0];
                    var headerCells = headerRow.Descendants<Cell>().OrderBy(cell => GetColumnIndex(cell.CellReference)).ToList();

                    int headerCurrentColumn = 0;
                    foreach (var cell in headerCells)
                    {
                        int cellColumn = GetColumnIndex(cell.CellReference);
                        // 处理空列
                        while (headerCurrentColumn < cellColumn)
                        {
                            columnNames.Add($"列 {headerCurrentColumn + 1}");
                            headerCurrentColumn++;
                        }
                        string cellValue = GetCellValue(cell, sharedStringTable);
                        columnNames.Add(cellValue);
                        headerCurrentColumn++;
                    }

                    // 读取数据行（包括所有行，不遗漏）
                    for (int i = 1; i < rows.Count; i++)
                    {
                        var row = rows[i];
                        var rowData = new List<string>();
                        var cells = row.Descendants<Cell>().OrderBy(cell => GetColumnIndex(cell.CellReference)).ToList();

                        int currentColumn = 0;
                        foreach (var cell in cells)
                        {
                            int cellColumn = GetColumnIndex(cell.CellReference);
                            // 处理空列，确保每一行的列数一致
                            while (currentColumn < cellColumn)
                            {
                                rowData.Add(string.Empty); // 添加空字符串而非null
                                currentColumn++;
                            }
                            string cellValue = GetCellValue(cell, sharedStringTable);
                            rowData.Add(cellValue);
                            currentColumn++;
                        }
                        result.Add(rowData);
                    }
                }
                UpdateStatus($"成功读取Excel数据，共 {result.Count} 行");
                return result;
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
                System.Diagnostics.Debug.WriteLine($"读取Excel数据异常: {ex.ToString()}");
                return null;
            }
        }

        // 生成参数值的核心方法
        private string GenerateParamValue(ParamConfig config, List<List<string>> excelData,
                                         int fileIndex, int paramBaseRow, int occurrenceIndex, out bool skipExcelRow)
        {
            skipExcelRow = false;
            string value = string.Empty;

            switch (config.ConfigType)
            {
                case "Excel列":
                    // 检查当前文件参数计数器中是否存在该参数
                    if (!_currentFileParamCounters.TryGetValue(config.ParamName, out int paramCounter))
                    {
                        paramCounter = 0; // 使用默认值
                        UpdateStatus($"警告：参数 {config.ParamName} 不在当前文件参数计数器中，使用默认值0");
                    }

                    int dataRowIndex;

                    if (config.ExcelMode == "整个文件使用同一行")
                    {
                        // 整个文件使用同一行 - 使用基于文件索引的起始行
                        dataRowIndex = paramBaseRow;
                    }
                    else // 顺序填写模式
                    {
                        // 使用全局位置追踪器获取当前应该使用的行索引
                        int globalPos = 0;
                        if (_excelParamGlobalPosition.TryGetValue(config.ParamName, out int existingPos))
                        {
                            globalPos = existingPos;
                        }
                        else
                        {
                            _excelParamGlobalPosition[config.ParamName] = 0;
                        }
                        dataRowIndex = globalPos;
                    }

                    // 检查数据行索引是否有效
                    if (dataRowIndex < 0 || dataRowIndex >= excelData.Count)
                    {
                        skipExcelRow = true;
                        UpdateStatus($"警告：参数 {config.ParamName} 的数据行索引 {dataRowIndex} 无效，已超出Excel数据范围");
                        return config.EmptyValue;
                    }

                    // 检查列索引是否有效
                    if (config.ExcelColumnIndex < 0 || config.ExcelColumnIndex >= excelData[dataRowIndex].Count)
                    {
                        skipExcelRow = true;
                        UpdateStatus($"警告：参数 {config.ParamName} 的列索引 {config.ExcelColumnIndex} 无效，已超出列范围");
                        return config.EmptyValue;
                    }

                    // 获取单元格值
                    value = excelData[dataRowIndex][config.ExcelColumnIndex];

                    // 精确判断空值：只有当值为null、空字符串或空白时才视为空
                    bool isEmptyValue = string.IsNullOrWhiteSpace(value);

                    if (isEmptyValue)
                    {
                        // 只有当用户设置了空值替换内容时才视为需要替换
                        skipExcelRow = !string.IsNullOrEmpty(config.EmptyValue);
                        if (skipExcelRow)
                        {
                            UpdateStatus($"参数 {config.ParamName} 在行 {dataRowIndex + 1} 列 {config.ExcelColumnIndex + 1} 的值为空，使用替换值");
                        }
                        else
                        {
                            // 如果没有设置替换值，使用原始空值
                            value = string.Empty;
                        }
                    }

                    // 对于顺序填写模式，每次使用后递增全局位置
                    if (config.ExcelMode == "顺序填写")
                    {
                        // 无论是否为空，都递增位置，确保顺序正确
                        int currentPos = _excelParamGlobalPosition[config.ParamName];
                        _excelParamGlobalPosition[config.ParamName] = currentPos + 1;
                        UpdateStatus($"参数 {config.ParamName}（顺序填写）使用行 {dataRowIndex + 1}，下次将使用行 {currentPos + 2}");
                    }
                    else if (config.ExcelMode != "整个文件使用同一行" && !skipExcelRow)
                    {
                        _currentFileParamCounters[config.ParamName] = paramCounter + 1;
                    }
                    break;

                case "定值":
                    value = config.FixedValue ?? string.Empty;
                    break;

                case "变值":
                    if (config.FixedValue.Contains("+"))
                    {
                        var parts = config.FixedValue.Split('+');
                        if (int.TryParse(parts[0], out int start) && int.TryParse(parts[1], out int step))
                        {
                            value = (start + config.GlobalSequence * step).ToString();
                            config.GlobalSequence++;
                        }
                        else
                        {
                            value = "格式错误";
                            UpdateStatus($"警告：参数 {config.ParamName} 的变值格式错误，应为 起始值+步长");
                        }
                    }
                    else
                    {
                        value = "格式错误";
                        UpdateStatus($"警告：参数 {config.ParamName} 的变值格式错误，应为 起始值+步长");
                    }
                    break;
            }

            return value ?? string.Empty;
        }
        #endregion
    }

    public class ParamConfig
    {
        public bool IsEnable { get; set; }
        public string ParamName { get; set; }
        public string ConfigType { get; set; }
        public string FixedValue { get; set; }
        public int ExcelColumnIndex { get; set; } = -1;
        public string ExcelMode { get; set; } = "整个文件使用同一行";
        public string EmptyValue { get; set; } = ""; // 默认空字符串而非null
        public int CurrentSequence { get; set; }
        public int GlobalSequence { get; set; }
        public int TotalRequiredRows { get; set; }
        public int RowOffsetPerFile { get; set; }

        public bool IsValid(out string errorMsg)
        {
            errorMsg = string.Empty;

            if (string.IsNullOrEmpty(ParamName))
            {
                errorMsg = "参数名不能为空";
                return false;
            }

            switch (ConfigType)
            {
                case "定值":
                    if (string.IsNullOrEmpty(FixedValue))
                    {
                        errorMsg = $"参数【{ParamName}】的定值不能为空";
                        return false;
                    }
                    break;

                case "变值":
                    if (string.IsNullOrEmpty(FixedValue) || !FixedValue.Contains("+"))
                    {
                        errorMsg = $"参数【{ParamName}】的变值规则格式错误，应为：起始值+步长";
                        return false;
                    }
                    break;

                case "Excel列":
                    if (ExcelColumnIndex < 0)
                    {
                        errorMsg = $"参数【{ParamName}】未选择Excel列";
                        return false;
                    }
                    break;

                default:
                    errorMsg = $"参数【{ParamName}】的配置类型未知";
                    return false;
            }

            return true;
        }

        public ParamConfig Clone()
        {
            return new ParamConfig
            {
                IsEnable = this.IsEnable,
                ParamName = this.ParamName,
                ConfigType = this.ConfigType,
                FixedValue = this.FixedValue,
                ExcelColumnIndex = this.ExcelColumnIndex,
                ExcelMode = this.ExcelMode,
                EmptyValue = this.EmptyValue,
                CurrentSequence = this.CurrentSequence,
                GlobalSequence = this.GlobalSequence,
                TotalRequiredRows = this.TotalRequiredRows,
                RowOffsetPerFile = this.RowOffsetPerFile
            };
        }
    }

    public class WatermarkAdorner : Adorner
    {
        private string _watermarkText;
        private Brush _watermarkBrush = Brushes.Gray;
        private Pen _borderPen;
        private double _pixelsPerDip;

        public WatermarkAdorner(UIElement adornedElement, string watermarkText)
            : base(adornedElement)
        {
            _watermarkText = watermarkText;
            _borderPen = new Pen(new SolidColorBrush(System.Windows.Media.Color.FromArgb(50, 0, 0, 0)), 1);
            IsHitTestVisible = false;
            adornedElement.GotFocus += AdornedElement_GotFocus;
            adornedElement.LostFocus += AdornedElement_LostFocus;

            _pixelsPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip;

            InvalidateVisual();
        }

        private void AdornedElement_GotFocus(object sender, RoutedEventArgs e)
        {
            InvalidateVisual();
        }

        private void AdornedElement_LostFocus(object sender, RoutedEventArgs e)
        {
            InvalidateVisual();
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);

            var textBox = AdornedElement as TextBox;
            if (textBox == null || !string.IsNullOrEmpty(textBox.Text) || textBox.IsFocused)
                return;

            drawingContext.DrawText(
                new FormattedText(
                    _watermarkText,
                    System.Globalization.CultureInfo.CurrentUICulture,
                    FlowDirection.LeftToRight,
                    new Typeface(textBox.FontFamily, textBox.FontStyle, textBox.FontWeight, textBox.FontStretch),
                    textBox.FontSize,
                    _watermarkBrush,
                    _pixelsPerDip),
                new Point(5, 2));
        }
    }
}
