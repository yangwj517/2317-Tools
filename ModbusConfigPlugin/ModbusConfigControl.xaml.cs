using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using ModbusConfigPlugin.pojo;
using ModbusConfigPlugin.pojo.Impl;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using Item = ModbusConfigPlugin.pojo.Item;

namespace ModbusConfigPlugin
{
    public partial class ModbusConfigControl : UserControl
    {
        private DataTable _excelData;
        private List<string> _availableColumns;

        public ModbusConfigControl()
        {
            InitializeComponent();
            InitializeControls();
        }

        private void InitializeControls()
        {
            // 初始化功能码选择
            DefaultFunctionCode.SelectedIndex = 5; // 默认选择05功能码

            // 初始化列选择框
            var columnComboBoxes = new[] {
                PointNameColumn,
                AddressColumn,DescriptionColumn,
                DeadbandColumn,OutputConversionColumn,
                AckStatusColumn,AckCommandColumn
            };

            foreach (var comboBox in columnComboBoxes)
            {
                comboBox.Items.Add("default");
                comboBox.SelectedIndex = 0;
            }
        }

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls",
                Title = "选择Excel文件"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                FilePathTextBox.Text = openFileDialog.FileName;
                LoadSheetNames(openFileDialog.FileName);
                UpdateStatus("Excel文件加载成功，请选择Sheet");
            }
        }

        private void LoadSheetNames(string filePath)
        {
            try
            {
                SheetComboBox.Items.Clear();

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = spreadsheet.WorkbookPart;
                    var sheets = workbookPart.Workbook.Sheets;

                    foreach (Sheet sheet in sheets)
                    {
                        SheetComboBox.Items.Add(sheet.Name);
                    }
                }

                if (SheetComboBox.Items.Count > 0)
                {
                    SheetComboBox.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"加载Sheet名称失败: {ex.Message}");
            }
        }

        private void LoadSheet_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FilePathTextBox.Text) || SheetComboBox.SelectedItem == null)
            {
                UpdateStatus("请先选择Excel文件和Sheet");
                return;
            }

            try
            {
                LoadExcelData(FilePathTextBox.Text, SheetComboBox.SelectedItem.ToString());
                UpdateColumnComboBoxes();
                UpdateStatus("Sheet数据加载成功，请配置列映射");
            }
            catch (Exception ex)
            {
                UpdateStatus($"加载Sheet数据失败: {ex.Message}");
            }
        }

        // 功能码选择后参数展示
        private void DefaultFunctionCode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DefaultFunctionCode.SelectedItem is ComboBoxItem selectedItem)
            {
                string functionCode = selectedItem.Tag?.ToString();
                UpdateDataTypeOptions(functionCode);

                // 隐藏所有面板
                RegisterMappingPanel.Visibility = Visibility.Collapsed;
                DataProcessingPanel.Visibility = Visibility.Collapsed;
                AlarmMappingPanel.Visibility = Visibility.Collapsed;
                ScannerMappingPanel.Visibility = Visibility.Collapsed;

                // 根据功能码显示相应的配置面板
                switch (functionCode)
                {
                    case "CODE_0":
                        ScannerMappingPanel.Visibility = Visibility.Visible;
                        AlarmMappingPanel.Visibility = Visibility.Visible;
                        break;

                    case "CODE_1": // 读取线圈
                    case "CODE_2": // 读取离散输入
                        ScannerMappingPanel.Visibility = Visibility.Visible;
                        break;

                    case "CODE_3": // 读取保持寄存器
                    case "CODE_4": // 读取输入寄存器
                        ScannerMappingPanel.Visibility = Visibility.Visible;
                        RegisterMappingPanel.Visibility = Visibility.Visible;
                        break;

                    case "CODE_6": // 写单个寄存器
                    case "CODE_16": // 写多个寄存器
                    case "CODE_22": // 屏蔽写寄存器
                        RegisterMappingPanel.Visibility = Visibility.Visible;
                        DataProcessingPanel.Visibility = Visibility.Visible;
                        break;

                    case "CODE_5": // 写单个线圈
                    case "CODE_15": // 写多个线圈
                               // 只显示基本配置
                        break;
                }
                UpdateStatus($"已选择功能码: {selectedItem.Content}");
            }
        }
        
        // 功能码选择后可选择数据类型更新
        private void UpdateDataTypeOptions(string functionCode)
        {
            // 清空现有选项
            DataTypeColumn.Items.Clear();

            // 根据功能码添加相应的数据类型选项
            switch (functionCode)
            {
                case "CODE_0": // Alarm With Ack
                case "CODE_1": // 01 - 读取线圈
                case "CODE_2": // 02 - 读取离散输入
                case "CODE_5": // 05 - 写单个线圈
                case "CODE_15": // 15 - 写多个线圈
                    DataTypeColumn.Items.Add("BOOL");
                    break;

                case "CODE_3": // 03 - 读取保持寄存器
                case "CODE_4": // 04 - 读取输入寄存器
                case "CODE_6": // 06 - 写单个寄存器
                case "CODE_16": // 16 - 写多个寄存器
                                // 寄存器功能码支持所有数据类型
                    DataTypeColumn.Items.Add("BYTE");
                    DataTypeColumn.Items.Add("UBYTE");
                    DataTypeColumn.Items.Add("INT16");
                    DataTypeColumn.Items.Add("UINT16");
                    DataTypeColumn.Items.Add("INT32");
                    DataTypeColumn.Items.Add("UINT32");
                    DataTypeColumn.Items.Add("FLOAT");
                    DataTypeColumn.Items.Add("BOOL");
                    break;
                case "CODE_22": // 22 - 屏蔽写寄存器
                                // 寄存器功能码支持所有数据类型
                    DataTypeColumn.Items.Add("BYTE");
                    DataTypeColumn.Items.Add("UBYTE");
                    DataTypeColumn.Items.Add("INT16");
                    DataTypeColumn.Items.Add("UINT16");
                    DataTypeColumn.Items.Add("BOOL");
                    break;
            }

            // 设置默认选择
            if (DataTypeColumn.Items.Count > 0)
            {
                DataTypeColumn.SelectedIndex = 0;
            }
        }

        private void LoadExcelData(string filePath, string sheetName)
        {
            _excelData = new DataTable();

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = spreadsheet.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);

                if (sheet == null)
                {
                    throw new Exception($"未找到Sheet: {sheetName}");
                }

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

                var rows = worksheetPart.Worksheet.Descendants<Row>().ToList();

                if (rows.Count == 0)
                {
                    throw new Exception("Sheet中没有数据");
                }

                // 获取列头（第一行）
                var headerRow = rows.First();
                foreach (Cell cell in headerRow)
                {
                    string columnName = GetCellValue(cell, sharedStringTable);
                    _excelData.Columns.Add(string.IsNullOrEmpty(columnName) ? $"Column{_excelData.Columns.Count + 1}" : columnName);
                }

                // 获取可用列列表
                _availableColumns = new List<string>();
                for (int i = 0; i < _excelData.Columns.Count; i++)
                {
                    _availableColumns.Add(_excelData.Columns[i].ColumnName);
                }

                // 读取数据行（从第二行开始）
                for (int i = 1; i < rows.Count; i++)
                {
                    var dataRow = _excelData.NewRow();
                    var row = rows[i];
                    var cells = row.Descendants<Cell>().ToList();

                    for (int j = 0; j < _excelData.Columns.Count; j++)
                    {
                        if (j < cells.Count)
                        {
                            dataRow[j] = GetCellValue(cells[j], sharedStringTable);
                        }
                        else
                        {
                            dataRow[j] = string.Empty;
                        }
                    }
                    _excelData.Rows.Add(dataRow);
                }
            }
        }

        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.CellValue == null) return string.Empty;

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sharedStringTable?.ElementAt(int.Parse(value))?.InnerText ?? value;
            }

            return value;
        }

        //更新列名
        private void UpdateColumnComboBoxes()
        {
            var columnComboBoxes = new[] {
                PointNameColumn,
                AddressColumn,DescriptionColumn,
                DeadbandColumn,OutputConversionColumn,
                AckStatusColumn,AckCommandColumn
            };

            foreach (var comboBox in columnComboBoxes)
            {
                // 保存当前选择
                var currentSelection = comboBox.SelectedItem?.ToString();
                comboBox.Items.Clear();
                comboBox.Items.Add("default");

                foreach (var column in _availableColumns)
                {
                    comboBox.Items.Add(column);
                }

                // 恢复选择
                if (!string.IsNullOrEmpty(currentSelection) && comboBox.Items.Contains(currentSelection))
                {
                    comboBox.SelectedItem = currentSelection;
                }
                else
                {
                    comboBox.SelectedIndex = 0;
                }
            }
        }

        private void SheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Sheet选择变化时的处理
        }

        private void PreviewData_Click(object sender, RoutedEventArgs e)
        {
            if (_excelData == null)
            {
                UpdateStatus("请先加载Excel数据");
                return;
            }

            PreviewDataGrid.ItemsSource = _excelData.DefaultView;
            UpdateStatus($"数据预览加载完成，共 {_excelData.Rows.Count} 行数据");
        }

        private void GenerateXml_Click(object sender, RoutedEventArgs e)
        {
            if (_excelData == null)
            {
                UpdateStatus("请先加载Excel数据");
                return;
            }

            try
            {
                GenerateModbusConfigXml();
            }
            catch (Exception ex)
            {
                UpdateStatus($"生成XML失败: {ex.Message}");
            }
        }

        private void GenerateModbusConfigXml()
        {
            // 1. 解析用户选择的列映射
            var columnMapping = GetColumnMapping();

            // 2. 根据功能码创建对应的Item对象列表
            var items = CreateItemsFromExcelData(columnMapping);

            // 3. 根据最大数量分组创建ScanBlocks
            var scanBlocks = CreateScanBlocks(items);

            // 4. 生成XML文件
            var xmlDoc = GenerateModbusConfigXml(scanBlocks);

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "XML文件 (*.xml)|*.xml",
                FileName = "ELC-ModbusConfig",
                DefaultExt = "xml"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                xmlDoc.Save(saveFileDialog.FileName);
                UpdateStatus($"XML配置文件已生成: {saveFileDialog.FileName}");
            }
        }

        // 解析用户选择的映射列
        private Dictionary<string, string> GetColumnMapping()
        {
            var mapping = new Dictionary<string, string>
            {
                ["PointName"] = PointNameColumn.SelectedIndex > 0 ? PointNameColumn.SelectedItem.ToString() : "default",
                ["Address"] = AddressColumn.SelectedIndex > 0 ? AddressColumn.SelectedItem.ToString() : "default",
                ["Description"] = DescriptionColumn.SelectedIndex > 0 ? DescriptionColumn.SelectedItem.ToString() : "default",
                ["Deadband"] = DeadbandColumn.SelectedIndex > 0 ? DeadbandColumn.SelectedItem.ToString() : "default",
                ["OutputConversion"] = OutputConversionColumn.SelectedIndex > 0 ? OutputConversionColumn.SelectedItem.ToString() : "default",
                ["AckStatus"] = AckStatusColumn.SelectedIndex > 0 ? AckStatusColumn.SelectedItem.ToString() : "default",
                ["AckCommand"] = AckCommandColumn.SelectedIndex > 0 ? AckCommandColumn.SelectedItem.ToString() : "default"
            };
            return mapping;
        }


        // 创建ScanBlock
        private List<ScanBlock> CreateScanBlocks(List<Item> items)
        {
            var scanBlocks = new List<ScanBlock>();

            if (!int.TryParse(GroupMaxNum.Text, out int  maxNum) || maxNum <= 0)
            {
                maxNum = 1000; // 默认值
            }

            var functionCode = ((ComboBoxItem)DefaultFunctionCode.SelectedItem)?.Tag?.ToString();
            string operation = GetOperationByFunctionCode(functionCode);


            // 对items按地址排序
            var sortedItems = items.OrderBy(item =>
            {
                if (int.TryParse(item.Address, out int address))
                    return address;
                return int.MaxValue;
            }).ToList();


            // 分组创建ScanBlock
            for (int i = 0; i < sortedItems.Count; i += maxNum)
            {
                var blockItems = sortedItems.Skip(i).Take(maxNum).ToList();
                var blockIndex = scanBlocks.Count + 1;

                var scanBlock = new ScanBlock(blockIndex, operation, maxNum);
                scanBlock.Items = blockItems;

                scanBlocks.Add(scanBlock);
            }

            return scanBlocks;
        }

        private string GetOperationByFunctionCode(string functionCode)
        {
            switch (functionCode)
            {
                case "CODE_1": return "Read Coil Status";
                case "CODE_2": return "Read Input Status";
                case "CODE_3": return "Read Holding Registers";
                case "CODE_4": return "Read Input Registers";
                case "CODE_5": return "Force Single Coil";
                case "CODE_6": return "Preset Single Register";
                case "CODE_15": return "Force Multiple Coils";
                case "CODE_16": return "Preset Multiple Registers";
                case "CODE_22": return "Mask Write Register";
                case "CODE_0": return "Read Alarm With Ack";
                default: return "Force Single Coil";
            }
        }

        // 创建Item
        private List<Item> CreateItemsFromExcelData(Dictionary<string, string> columnMapping)
        {
            var items = new List<Item>();
            var functionCode = ((ComboBoxItem)DefaultFunctionCode.SelectedItem)?.Tag?.ToString();

            foreach (DataRow row in _excelData.Rows)
            {
                Item item = null;

                // 获取基本数据
                string pointName = GetRowValue(row, columnMapping["PointName"]);
                string dataType = DataTypeColumn.SelectedItem?.ToString();
                string address = GetRowValue(row, columnMapping["Address"]);
                string description = GetRowValue(row, columnMapping["Description"]);
                string scadaScanner = (bool)ScadaScannerCheckBox.IsChecked ? "Yes" : "No";
                string ackStatus = GetRowValue(row, columnMapping["AckStatus"]);
                string ackCommand = GetRowValue(row, columnMapping["AckCommand"]);
                string swapBytes = (bool)ByteReverseCheckBox.IsChecked ? "Yes" : "No";
                string swapWords = (bool)WordReverseCheckBox.IsChecked ? "Yes" : "No";
                string deadband = GetRowValue(row, columnMapping["Deadband"]);
                string outputConversion = columnMapping["OutputConversion"].Equals("default") ? "None" : GetRowValue(row, columnMapping["OutputConversion"]);

                // 根据功能码创建对应的Item对象
                switch (functionCode)
                {
                    case "CODE_0":
                        {
                            item = new FunCode_Alarm(dataType,pointName,address,description,scadaScanner,
                                ackStatus,ackCommand);
                            break;
                        }
                    case "CODE_1":
                        { 
                            item = new FunCode_1(pointName,address,description,dataType,scadaScanner);
                            break;
                        }
                    case "CODE_2":
                        {
                            item = new FunCode_2(pointName, address, description, dataType, scadaScanner);
                            break;
                        }                  
                    case "CODE_3":
                        {
                            item = new FunCode_3(pointName, address, description, dataType,
                                               scadaScanner, swapBytes, swapWords);
                            break;
                        }
                    case "CODE_4":
                        {
                            item = new FunCode_4(pointName, address, description, dataType,
                                               scadaScanner, swapBytes, swapWords);
                            break;
                        }
                    case "CODE_5":
                        {
                            item = new FunCode_5(pointName, address, description, dataType);
                            break;
                        }
                    case "CODE_6":
                        {
                            item = new FunCode_6(pointName, address, description, dataType,swapBytes,
                                swapWords,deadband,outputConversion);
                            break;
                        }
                    case "CODE_15":
                        {
                            item = new FunCode_15(pointName, address, description, dataType);
                            break;
                        }
                    case "CODE_16":
                        {
                            item = new FunCode_16(pointName, address, description, dataType, swapBytes, 
                                swapWords, deadband, outputConversion);
                            break;
                        }
                    case "CODE_22":
                        {
                            item = new FunCode_22(pointName, address, description, dataType, swapBytes, 
                                swapWords, deadband, outputConversion);
                            break;
                        }
                    default:
                        UpdateStatus("item选择时异常啦。。。");
                        break;
                }

                if (item != null)
                {
                    items.Add(item);
                }
            }

            return items;
        }



        //sheet数据获取
        private string GetRowValue(DataRow row, string columnName)
        {
            if (string.IsNullOrEmpty(columnName) || columnName == "default")
                return "";

            return row[columnName]?.ToString() ?? "";
        }

        private XDocument GenerateModbusConfigXml(List<ScanBlock> scanBlocks)
        {
            var root = new XElement("ELC-ModbusConfig");

            foreach (var scanBlock in scanBlocks)
            {
                root.Add(scanBlock.ToXElement());
            }

            return new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                root
            );
        }


        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            FilePathTextBox.Text = string.Empty;
            SheetComboBox.Items.Clear();
            _excelData = null;
            PreviewDataGrid.ItemsSource = null;
            InitializeControls();
            UpdateStatus("已清空所有配置");
        }

        private void UpdateStatus(string message)
        {
            StatusText.Text = $"[{DateTime.Now:HH:mm:ss}] {message}";
        }
    }
}