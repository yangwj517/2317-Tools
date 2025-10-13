using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Oxs = DocumentFormat.OpenXml.Spreadsheet; // 为Spreadsheet命名空间添加别名
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace BatchStats
{
    public partial class BatchStatsControl : UserControl
    {
        private string _folderPath;
        private uint _nextSheetId = 1;

        public BatchStatsControl()
        {
            InitializeComponent();
            AddLogMessage("就绪 - 请选择文件夹并开始生成");
        }

        private void AddLogMessage(string message, string type = "INFO")
        {
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                string logEntry = $"[{timestamp}] [{type}] {message}";
                LogListBox.Items.Add(logEntry);

                if (LogListBox.Items.Count > 0)
                {
                    LogListBox.ScrollIntoView(LogListBox.Items[LogListBox.Items.Count - 1]);
                }
            });
        }

        private void SelectFlord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new OpenFileDialog
                {
                    ValidateNames = false,
                    CheckFileExists = false,
                    CheckPathExists = true,
                    FileName = "选择文件夹",
                    Title = "请选择包含 .stats 文件的文件夹"
                };

                if (!string.IsNullOrEmpty(FlordPathTextBox.Text) && Directory.Exists(FlordPathTextBox.Text))
                {
                    dialog.InitialDirectory = FlordPathTextBox.Text;
                }
                else
                {
                    dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }

                if (dialog.ShowDialog() == true)
                {
                    _folderPath = Path.GetDirectoryName(dialog.FileName);
                    FlordPathTextBox.Text = _folderPath;
                    AddLogMessage($"已选择文件夹: {_folderPath}");
                    CheckStatsFiles(_folderPath);
                }
                else
                {
                    AddLogMessage("用户取消了文件夹选择");
                }
            }
            catch (Exception ex)
            {
                AddLogMessage($"选择文件夹时出错: {ex.Message}", "ERROR");
                MessageBox.Show($"选择文件夹时出错: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CheckStatsFiles(string folderPath)
        {
            try
            {
                string[] statsFiles = Directory.GetFiles(folderPath, "*.stats", SearchOption.TopDirectoryOnly);
                AddLogMessage($"找到 {statsFiles.Length} 个 .stats 文件");

                if (statsFiles.Length > 0)
                {
                    for (int i = 0; i < Math.Min(3, statsFiles.Length); i++)
                    {
                        AddLogMessage($"  - {Path.GetFileName(statsFiles[i])}");
                    }
                    if (statsFiles.Length > 3)
                    {
                        AddLogMessage($"  - ... 还有 {statsFiles.Length - 3} 个文件");
                    }
                }
                else
                {
                    AddLogMessage("警告: 未找到 .stats 文件", "WARNING");
                }
            }
            catch (Exception ex)
            {
                AddLogMessage($"检查文件时出错: {ex.Message}", "ERROR");
            }
        }

        private bool ParseStatsFile(string filePath, out List<List<string>> tableData, out List<string> tableHeaders)
        {
            tableData = new List<List<string>>();
            tableHeaders = new List<string>();
            bool isTableDataStarted = false;
            bool isHeaderParsed = false;
            string fileName = Path.GetFileName(filePath);

            try
            {
                string fileContent = File.ReadAllText(filePath, DetectFileEncoding(filePath));
                string[] lines = fileContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();

                    if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith("#"))
                    {
                        if (trimmedLine.StartsWith("# ColHeaders "))
                        {
                            isTableDataStarted = true;
                            string headerStr = trimmedLine.Substring("# ColHeaders ".Length).Trim();
                            tableHeaders = Regex.Split(headerStr, @"\s+")
                                             .Where(h => !string.IsNullOrWhiteSpace(h))
                                             .ToList();
                            isHeaderParsed = true;
                        }
                        continue;
                    }

                    if (isTableDataStarted)
                    {
                        if (!isHeaderParsed)
                        {
                            tableHeaders = Regex.Split(trimmedLine, @"\s+")
                                             .Where(h => !string.IsNullOrWhiteSpace(h))
                                             .ToList();
                            isHeaderParsed = true;
                            continue;
                        }

                        List<string> rowData = Regex.Split(trimmedLine, @"\s+")
                                               .Where(cell => !string.IsNullOrWhiteSpace(cell))
                                               .ToList();

                        if (rowData.Count == tableHeaders.Count)
                        {
                            tableData.Add(rowData);
                        }
                        else
                        {
                            AddLogMessage($"[{fileName}] 跳过异常行（{rowData.Count}列/预期{tableHeaders.Count}列）：{trimmedLine}", "WARNING");
                        }
                    }
                }

                if (tableData.Count == 0 || tableHeaders.Count == 0)
                {
                    AddLogMessage($"[{fileName}] 解析失败：未找到有效表格数据", "ERROR");
                    return false;
                }

                AddLogMessage($"[{fileName}] 解析成功（表格{tableData.Count}行，列头{tableHeaders.Count}个）", "SUCCESS");
                return true;
            }
            catch (Exception ex)
            {
                AddLogMessage($"[{fileName}] 解析出错：{ex.Message}", "ERROR");
                return false;
            }
        }

        private Encoding DetectFileEncoding(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                byte[] bom = new byte[3];
                int readLen = fs.Read(bom, 0, 3);
                if (readLen >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
                {
                    return Encoding.UTF8;
                }

                fs.Position = 0;
                while (fs.ReadByte() != -1)
                {
                    if (fs.Position < fs.Length)
                    {
                        byte b = (byte)fs.ReadByte();
                        if (b > 127)
                        {
                            return Encoding.UTF8;
                        }
                    }
                }
                return Encoding.ASCII;
            }
        }

        private void GenerateXlsx_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_folderPath) || !Directory.Exists(_folderPath))
            {
                AddLogMessage("请先选择有效的文件夹", "WARNING");
                MessageBox.Show("请先选择包含.stats文件的文件夹", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string[] statsFiles = Directory.GetFiles(_folderPath, "*.stats", SearchOption.TopDirectoryOnly);
            if (statsFiles.Length == 0)
            {
                AddLogMessage("当前文件夹无.stats文件，无法生成XLSX", "WARNING");
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "Excel文件 (*.xlsx)|*.xlsx",
                Title = "保存XLSX文件",
                FileName = $"Stats_Export_{DateTime.Now:yyyyMMddHHmmss}.xlsx",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                OverwritePrompt = true
            };

            if (saveDialog.ShowDialog() != true)
            {
                AddLogMessage("用户取消了XLSX保存", "INFO");
                return;
            }
            string xlsxPath = saveDialog.FileName;
            AddLogMessage("开始生成XLSX文件...");

            try
            {
                _nextSheetId = 1;

                using (SpreadsheetDocument document = SpreadsheetDocument.Create(xlsxPath, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Oxs.Workbook();

                    // 添加样式表
                    WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = CreateXlsxStylesheet();
                    stylesPart.Stylesheet.Save();

                    // 创建Sheets集合
                    Oxs.Sheets sheets = workbookPart.Workbook.AppendChild(new Oxs.Sheets());

                    foreach (string statsFile in statsFiles)
                    {
                        string sheetName = Path.GetFileNameWithoutExtension(statsFile);
                        if (sheetName.Length > 31)
                        {
                            sheetName = sheetName.Substring(0, 28) + "...";
                            AddLogMessage($"文件 {Path.GetFileName(statsFile)} 名称过长，Sheet名截断为: {sheetName}", "WARNING");
                        }

                        if (!ParseStatsFile(statsFile, out var tableData, out var tableHeaders))
                        {
                            AddLogMessage($"跳过文件: {Path.GetFileName(statsFile)}", "WARNING");
                            continue;
                        }

                        // 创建工作表
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        Oxs.Worksheet worksheet = new Oxs.Worksheet();
                        Oxs.SheetData sheetData = new Oxs.SheetData();

                        // 添加列宽定义
                        Oxs.Columns columns = new Oxs.Columns();
                        for (int i = 0; i < tableHeaders.Count; i++)
                        {
                            double width = Math.Max(tableHeaders[i].Length * 1.2, 10);
                            columns.Append(new Oxs.Column
                            {
                                Min = (uint)(i + 1),
                                Max = (uint)(i + 1),
                                Width = width,
                                CustomWidth = true
                            });
                        }
                        worksheet.Append(columns);
                        worksheet.Append(sheetData);

                        // 写入数据
                        int currentRow = 1;
                        if (tableData.Count > 0 && tableHeaders.Count > 0)
                        {
                            AddXlsxRow(sheetData, currentRow++, tableHeaders, isHeader: true);
                            foreach (var rowData in tableData)
                            {
                                AddXlsxRow(sheetData, currentRow++, rowData);
                            }
                        }

                        worksheetPart.Worksheet = worksheet;
                        worksheetPart.Worksheet.Save();

                        // 添加工作表到工作簿
                        Oxs.Sheet sheet = new Oxs.Sheet
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            Name = sheetName,
                            SheetId = _nextSheetId++
                        };
                        sheets.Append(sheet);

                        AddLogMessage($"已生成Sheet: {sheetName}", "SUCCESS");
                    }

                    // 添加基础文档范围定义
                    workbookPart.Workbook.AppendChild(new Oxs.SheetDimension { Reference = "A1" });
                    workbookPart.Workbook.Save();
                }

                Dispatcher.Invoke(() =>
                {
                    AddLogMessage($"XLSX文件生成完成！路径: {xlsxPath}", "SUCCESS");
                    MessageBox.Show($"XLSX文件生成成功！\n路径: {xlsxPath}", "成功",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    AddLogMessage($"生成XLSX文件出错: {ex.Message}", "ERROR");
                    MessageBox.Show($"生成XLSX文件失败: {ex.Message}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private Oxs.Stylesheet CreateXlsxStylesheet()
        {
            // 字体定义
            Oxs.Fonts fonts = new Oxs.Fonts();
            fonts.Append(new Oxs.Font
            {
                FontName = new Oxs.FontName { Val = "Arial" },
                FontSize = new Oxs.FontSize { Val = 10 }
            });
            fonts.Append(new Oxs.Font
            {
                FontName = new Oxs.FontName { Val = "Arial" },
                FontSize = new Oxs.FontSize { Val = 10 },
                Bold = new Oxs.Bold()
            });
            fonts.Count = (uint)fonts.ChildElements.Count;

            // 填充定义
            Oxs.Fills fills = new Oxs.Fills();
            fills.Append(new Oxs.Fill
            {
                PatternFill = new Oxs.PatternFill { PatternType = Oxs.PatternValues.None }
            });
            fills.Append(new Oxs.Fill
            {
                PatternFill = new Oxs.PatternFill { PatternType = Oxs.PatternValues.Gray125 }
            });
            fills.Append(new Oxs.Fill
            {
                PatternFill = new Oxs.PatternFill
                {
                    PatternType = Oxs.PatternValues.Solid,
                    ForegroundColor = new Oxs.ForegroundColor { Rgb = "E6F3FF" },
                    BackgroundColor = new Oxs.BackgroundColor { Indexed = 64 }
                }
            });
            fills.Count = (uint)fills.ChildElements.Count;

            // 边框定义
            Oxs.Borders borders = new Oxs.Borders();
            Oxs.Border border = new Oxs.Border
            {
                LeftBorder = new Oxs.LeftBorder
                {
                    Style = Oxs.BorderStyleValues.Thin,
                    Color = new Oxs.Color { Indexed = 64 }
                },
                RightBorder = new Oxs.RightBorder
                {
                    Style = Oxs.BorderStyleValues.Thin,
                    Color = new Oxs.Color { Indexed = 64 }
                },
                TopBorder = new Oxs.TopBorder
                {
                    Style = Oxs.BorderStyleValues.Thin,
                    Color = new Oxs.Color { Indexed = 64 }
                },
                BottomBorder = new Oxs.BottomBorder
                {
                    Style = Oxs.BorderStyleValues.Thin,
                    Color = new Oxs.Color { Indexed = 64 }
                },
                DiagonalBorder = new Oxs.DiagonalBorder()
            };
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            // 单元格格式
            Oxs.CellFormats cellFormats = new Oxs.CellFormats();
            cellFormats.Append(new Oxs.CellFormat
            {
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyBorder = true
            });
            cellFormats.Append(new Oxs.CellFormat
            {
                FontId = 1,
                FillId = 2,
                BorderId = 0,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true
            });
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            // 组装样式表
            Oxs.Stylesheet stylesheet = new Oxs.Stylesheet();
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellFormats);

            return stylesheet;
        }

        private void AddXlsxRow(Oxs.SheetData sheetData, int rowIndex, List<string> cellValues, bool isHeader = false)
        {
            Oxs.Row row = new Oxs.Row { RowIndex = (uint)rowIndex };

            for (int colIndex = 0; colIndex < cellValues.Count; colIndex++)
            {
                string colName = GetXlsxColumnName(colIndex + 1);
                string cellRef = $"{colName}{rowIndex}";

                Oxs.Cell cell = new Oxs.Cell
                {
                    CellReference = cellRef,
                    CellValue = new Oxs.CellValue(cellValues[colIndex]),
                    StyleIndex = isHeader ? 1U : 0U
                };

                // 设置单元格数据类型
                if (double.TryParse(cellValues[colIndex], out _) ||
                    decimal.TryParse(cellValues[colIndex], out _))
                {
                    cell.DataType = Oxs.CellValues.Number;
                }
                else if (DateTime.TryParse(cellValues[colIndex], out _))
                {
                    cell.DataType = Oxs.CellValues.Date;
                }
                else
                {
                    cell.DataType = Oxs.CellValues.String;
                }

                row.Append(cell);
            }

            sheetData.Append(row);
        }

        private string GetXlsxColumnName(int columnIndex)
        {
            string columnName = string.Empty;
            while (columnIndex > 0)
            {
                columnIndex--;
                char colChar = (char)('A' + columnIndex % 26);
                columnName = colChar + columnName;
                columnIndex /= 26;
            }
            return columnName;
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                FlordPathTextBox.Text = string.Empty;
                LogListBox.Items.Clear();
                _folderPath = string.Empty;
                AddLogMessage("已清空所有选择和日志", "INFO");
            });
        }
    }
}
