using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Toolbox.Core;

namespace ToolboxHost
{
    public partial class MainWindow : Window
    {
        private PluginManager _pluginManager;

        public MainWindow()
        {
            // 初始化日志系统
            Logger.Initialize();
            Logger.Info("应用程序启动");
            InitializeComponent();
            InitializeApplication();
        }

        private void InitializeApplication()
        {
            try
            {
                Logger.Info("开始初始化应用程序");
                // 初始化插件管理器
                _pluginManager = new PluginManager();

                // 订阅插件事件
                _pluginManager.PluginLoaded += OnPluginLoaded;
                _pluginManager.PluginUnloaded += OnPluginUnloaded;

                // 加载插件
                _pluginManager.LoadAllPlugins();

                // 更新插件列表
                UpdatePluginList();
                UpdateStatus("应用程序初始化完成");

                Logger.Info("应用程序初始化完成");
            }
            catch (Exception ex)
            {
                Logger.Error("应用程序初始化失败", ex);
                MessageBox.Show($"应用程序初始化失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// 新插件导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportNewPlugin_Click(object sender, RoutedEventArgs e)
        {
            // 用户选择一个文件夹
            // 1. 让用户选择插件所在的文件夹
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "请选择插件所在的文件夹",
                ShowNewFolderButton = false // 不允许新建文件夹
            };

            if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return; // 用户取消选择
            }
            // 检测文件夹结构是否符合插件要求
            string sourceDir = folderDialog.SelectedPath;// 获取文件夹地址
            string selectedFolderName = Path.GetFileName(sourceDir);  // 获取文件夹名称
            string pluginDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins");
            try
            {
                // 结构是否符合要求
                if (!IsValidPluginStructure(sourceDir,selectedFolderName))
                {
                    MessageBox.Show("插件文件夹结构不符合要求！\n要求：可创建插件模板查看具体插件文件结构要求。",
                         "导入失败",
                         MessageBoxButton.OK,
                         MessageBoxImage.Error);
                    return;
                }
                // 是否存在同名插件
                if (hasSamePlugin(pluginDir,selectedFolderName))
                {
                    MessageBoxResult result = MessageBox.Show("此插件目录已在系统中存在\n是否覆盖？",
                         "插件已存在！",
                         MessageBoxButton.OKCancel,
                         MessageBoxImage.Warning);

                    if (result == MessageBoxResult.OK)
                    {
                        // 执行覆盖操作
                        // 删除原有插件
                        DelFlord(pluginDir,selectedFolderName);
                    }
                    else
                    {
                        // 不执行操作,退出插件导入
                        return;
                    }
                }
                // 导入该插件
                ImportPlugin(sourceDir, Path.Combine(pluginDir, selectedFolderName));
                // 刷新插件列表
                _pluginManager.ReloadAllPlugins();
                WorkspaceTabControl.Items.Clear();
                MessageBox.Show("新插件导入成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                Logger.Info($"导入新插件{selectedFolderName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("新插件导入异常！", "失败", MessageBoxButton.OK, MessageBoxImage.Error);
                Logger.Error($"插件导入异常：{ex.Message}", ex);
            }
            
        }

        /// <summary>
        /// 复制文件夹及其内容（增强版，解决文件锁定问题）
        /// </summary>
        private void ImportPlugin(string sourceDir, string destDir)
        {
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // 复制文件（带重试和解锁逻辑）
            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string destFile = Path.Combine(destDir, Path.GetFileName(file));
                // 使用增强的文件复制方法
                CopyFileWithRetry(file, destFile, true);
            }

            // 递归复制子文件夹
            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                string destSubDir = Path.Combine(destDir, Path.GetFileName(subDir));
                ImportPlugin(subDir, destSubDir);
            }
        }

        /// <summary>
        /// 带重试和解锁逻辑的文件复制
        /// </summary>
        private bool CopyFileWithRetry(string sourcePath, string destPath, bool overwrite, int maxRetries = 5)
        {
            int retryCount = 0;
            while (retryCount < maxRetries)
            {
                try
                {
                    // 尝试释放目标文件可能的锁定
                    if (File.Exists(destPath))
                    {
                        ReleaseFileLock(destPath);
                    }

                    File.Copy(sourcePath, destPath, overwrite);
                    return true;
                }
                catch (IOException ex) when (IsFileLocked(ex))
                {
                    retryCount++;
                    if (retryCount >= maxRetries)
                    {
                        Logger.Error($"复制文件失败，已达最大重试次数: {sourcePath}", ex);
                        return false;
                    }

                    // 指数退避重试（100ms, 200ms, 400ms...）
                    int delay = (int)Math.Pow(2, retryCount) * 100;
                    Logger.Warning($"文件被锁定，将在 {delay}ms 后重试: {sourcePath}");
                    System.Threading.Thread.Sleep(delay);
                }
                catch (Exception ex)
                {
                    Logger.Error($"复制文件时发生意外错误: {sourcePath}", ex);
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// 尝试释放文件锁定
        /// </summary>
        private void ReleaseFileLock(string filePath)
        {
            try
            {
                // 检查文件是否被锁定
                if (IsFileLocked(filePath))
                {
                    Logger.Info($"尝试释放文件锁定: {filePath}");

                    // 方法1: 尝试以只读方式打开并立即关闭
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Delete))
                    {
                        stream.Close();
                    }

                    // 方法2: 强制垃圾回收，可能释放文件句柄
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch
            {
                // 释放失败时不抛出异常，继续重试
            }
        }

        /// <summary>
        /// 检查文件是否被锁定
        /// </summary>
        private bool IsFileLocked(Exception ex)
        {
            int errorCode = System.Runtime.InteropServices.Marshal.GetHRForException(ex) & ((1 << 16) - 1);
            return errorCode == 32 || errorCode == 33;
        }

        /// <summary>
        /// 检查文件是否被锁定
        /// </summary>
        private bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                using (File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException ex)
            {
                return IsFileLocked(ex);
            }
        }


        // 同名插件卸载
        private bool DelFlord(string pluginDir,string flordName)
        {
            try
            {
                IPlugin plugin = null;
                foreach (IPlugin x in _pluginManager.Plugins)
                {
                    if (x.ToolName == flordName)
                    {
                        plugin = x;
                        break;
                    }
                }
                if (plugin == null) return false;
                // 卸载该插件
                UnloadPlugin(plugin);
                // 删除文件夹
                Directory.Delete(Path.Combine(pluginDir,flordName), true);
            }
            catch (Exception ex) {
                Logger.Error("重名插件卸载失败！");
                return false;
            }   
            return true;
        }

        private bool hasSamePlugin(string pluginDir ,string flordName)
        {
            return Directory.Exists(Path.Combine(pluginDir, flordName));
        }



        // 插件目录是否符合设计
        private bool IsValidPluginStructure(string folderPath,string flordName)
        {
            // 检测所选文件夹下是否有且仅有一个Dll主程序，并且这个文件的名称和插件根目录文件夹名称一致
            string searchPattern = $"*.dll";
            string[] files = Directory.GetFiles(folderPath, searchPattern, SearchOption.TopDirectoryOnly);
            if (files.Length != 1) return false; 
            string fileName = Path.GetFileNameWithoutExtension(files[0]);
            if(flordName != fileName) return false;
            // 检测所选文件夹下是否存下依赖文件夹、资源文件夹
            bool hasDependenciesFlord = Directory.Exists(Path.Combine(folderPath, "Dependencies"));
            bool hasResourcesFlord = Directory.Exists(Path.Combine(folderPath, "Resources"));
            // 检查是否存在插件配置文件、介绍文件
            bool hasConfigFile = File.Exists(Path.Combine(folderPath, "plugin.json"));
            bool hasReadMeFile = File.Exists(Path.Combine(folderPath, "README.txt"));
            return hasConfigFile&&hasDependenciesFlord&&hasReadMeFile&&hasResourcesFlord;
        }


        private void UpdatePluginList()
        {
            Dispatcher.Invoke(() =>
            {
                PluginListBox.ItemsSource = null;
                PluginListBox.ItemsSource = _pluginManager.Plugins;
                PluginCountText.Text = $"插件: {_pluginManager.Plugins.Count}";
            });
        }


        private void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                StatusText.Text = message;
            });
            Logger.Info($"状态更新: {message}");
        }

        #region 事件处理方法

        private void OnPluginLoaded(IPlugin plugin, PluginInfo pluginInfo)
        {
            Dispatcher.Invoke(() =>
            {
                UpdatePluginList();
                UpdateStatus($"插件加载成功: {plugin.Name}");
            });

            Logger.Info($"插件加载事件: {plugin.Name}");
        }

        private void OnPluginUnloaded(IPlugin plugin, PluginInfo pluginInfo)
        {
            Dispatcher.Invoke(() =>
            {
                UpdatePluginList();
                UpdateStatus($"插件已卸载: {plugin.Name}");
            });

            Logger.Info($"插件卸载事件: {plugin.Name}");
        }

        private void PluginListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PluginListBox.SelectedItem is IPlugin selectedPlugin)
            {
                OpenPluginInTab(selectedPlugin);
            }
        }

        private void WorkspaceTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (WorkspaceTabControl.SelectedItem is TabItem selectedTab && selectedTab.Tag is IPlugin plugin)
            {
                WorkspaceTitle.Text = $"正在使用: {plugin.Name}";
            }
            else
            {
                WorkspaceTitle.Text = "请选择一个插件开始使用";
            }
        }

        private void TogglePluginList_Checked(object sender, RoutedEventArgs e)
        {
            if (PluginListColumn != null)
            {
                PluginListColumn.Width = new GridLength(300);
            }
        }

        private void TogglePluginList_Unchecked(object sender, RoutedEventArgs e)
        {
            if (PluginListColumn != null)
            {
                PluginListColumn.Width = new GridLength(0);
            }
        }

        private void ToggleStatusBar_Checked(object sender, RoutedEventArgs e)
        {
            if (StatusBar != null)
            {
                StatusBar.Visibility = Visibility.Visible;
            }
        }

        private void ToggleStatusBar_Unchecked(object sender, RoutedEventArgs e)
        {
            if (StatusBar != null)
            {
                StatusBar.Visibility = Visibility.Collapsed;
            }
        }

        #endregion

        #region 插件管理功能

        /// <summary>
        /// 打开插件
        /// </summary>
        private void OpenPluginInTab(IPlugin plugin)
        {
            // 检查是否已经打开了该插件的标签页
            foreach (TabItem tab in WorkspaceTabControl.Items)
            {
                if (tab.Tag == plugin)
                {
                    WorkspaceTabControl.SelectedItem = tab;
                    return;
                }
            }

            try
            {
                // 创建新的标签页
                var pluginControl = plugin.GetControl();
                if (pluginControl != null)
                {
                    TabItem newTab = new TabItem
                    {
                        Header = CreateTabHeader(plugin),
                        Content = pluginControl,
                        Tag = plugin
                    };

                    WorkspaceTabControl.Items.Add(newTab);
                    WorkspaceTabControl.SelectedItem = newTab;
                    WorkspaceTitle.Text = $"正在使用: {plugin.Name}";

                    UpdateStatus($"已打开插件: {plugin.Name}");
                    Logger.Info($"打开插件: {plugin.Name}");
                }
                else
                {
                    MessageBox.Show($"插件 {plugin.Name} 返回的界面为空", "警告",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"打开插件失败: {plugin.Name}", ex);
                MessageBox.Show($"打开插件失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 创建标签页标题（包含关闭按钮）
        /// </summary>
        private object CreateTabHeader(IPlugin plugin)
        {
            var stackPanel = new StackPanel { Orientation = Orientation.Horizontal };

            // 插件名称
            var textBlock = new TextBlock
            {
                Text = $"{plugin.Name} v{plugin.Version}",
                Margin = new Thickness(0, 0, 5, 0)
            };

            // 关闭按钮
            var closeButton = new Button
            {
                Content = "×",
                FontWeight = FontWeights.Bold,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(3, 0, 3, 0),
                Margin = new Thickness(0),
                Cursor = Cursors.Hand
            };

            closeButton.Click += (s, e) =>
            {
                // 找到对应的标签页并关闭
                var tab = WorkspaceTabControl.Items
                    .Cast<TabItem>()
                    .FirstOrDefault(t => t.Tag == plugin);

                if (tab != null)
                {
                    WorkspaceTabControl.Items.Remove(tab);

                    // 如果没有标签页了，更新标题
                    if (WorkspaceTabControl.Items.Count == 0)
                    {
                        WorkspaceTitle.Text = "请选择一个插件开始使用";
                    }
                }
            };

            stackPanel.Children.Add(textBlock);
            stackPanel.Children.Add(closeButton);

            return stackPanel;
        }

        /// <summary>
        /// 插件列表双击事件 - 打开插件
        /// </summary>
        private void PluginListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (PluginListBox.SelectedItem is IPlugin selectedPlugin)
            {
                OpenPluginInTab(selectedPlugin);
            }
        }

        #endregion

        #region 插件右键菜单功能

        /// <summary>
        /// 右键菜单 - 打开选中的插件
        /// </summary>
        private void OpenSelectedPlugin_Click(object sender, RoutedEventArgs e)
        {
            if (PluginListBox.SelectedItem is IPlugin selectedPlugin)
            {
                OpenPluginInTab(selectedPlugin);
            }
        }

        /// <summary>
        /// 打开插件文件夹
        /// </summary>
        private void OpenSelectedPluginFolder_Click(object sender, RoutedEventArgs e)
        {
            if (PluginListBox.SelectedItem is IPlugin selectedPlugin)
            {
                OpenPluginFolder(selectedPlugin);
            }
        }

        /// <summary>
        /// 查看插件信息（JSON格式）
        /// </summary>
        private void ViewPluginInfo_Click(object sender, RoutedEventArgs e)
        {
            if (PluginListBox.SelectedItem is IPlugin selectedPlugin)
            {
                ShowPluginInfo(selectedPlugin);
            }
            else
            {
                MessageBox.Show("请先选择一个插件", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// 卸载选中插件
        /// </summary>
        private void UnloadSelectedPlugin_Click(object sender, RoutedEventArgs e)
        {
            if (PluginListBox.SelectedItem is IPlugin selectedPlugin)
            {
                UnloadPlugin(selectedPlugin);
            }
            else
            {
                MessageBox.Show("请先选择一个插件", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// 显示插件信息（JSON格式）
        /// </summary>
        private void ShowPluginInfo(IPlugin plugin)
        {
            try
            {
                var config = _pluginManager.GetPluginConfig(plugin);

                // 创建JSON显示窗口
                var infoWindow = new PluginInfoWindow(plugin, config);
                infoWindow.Owner = this;
                infoWindow.ShowDialog();

                Logger.Info($"查看插件信息: {plugin.Name}");
            }
            catch (Exception ex)
            {
                Logger.Error($"显示插件信息失败: {plugin.Name}", ex);
                MessageBox.Show($"显示插件信息失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 打开插件文件夹
        /// </summary>
        private void OpenPluginFolder(IPlugin plugin)
        {
            try
            {
                var loadContext = _pluginManager.GetPluginLoadContext(plugin);
                if (loadContext != null && System.IO.Directory.Exists(loadContext.FolderPath))
                {
                    System.Diagnostics.Process.Start("explorer.exe", loadContext.FolderPath);
                    UpdateStatus($"已打开插件文件夹: {plugin.Name}");
                    Logger.Info($"打开插件文件夹: {plugin.Name}");
                }
                else
                {
                    MessageBox.Show("未找到插件的文件夹", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"打开插件文件夹失败: {plugin.Name}", ex);
                MessageBox.Show($"打开插件文件夹失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region 插件卸载功能

        /// <summary>
        /// 卸载指定插件
        /// </summary>
        private void UnloadPlugin(IPlugin plugin)
        {
            try
            {
                Logger.Info($"用户请求卸载插件: {plugin.Name}");

                // 确认对话框
                var result = MessageBox.Show($"确定要卸载插件 '{plugin.Name}' 吗？", "确认卸载",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    // 关闭该插件的所有标签页
                    ClosePluginTabs(plugin);

                    // 卸载插件
                    bool success = _pluginManager.UnloadPlugin(plugin);

                    if (success)
                    {
                        UpdateStatus($"插件已卸载: {plugin.Name}");
                        Logger.Info($"插件卸载成功: {plugin.Name}");
                    }
                    else
                    {
                        Logger.Warning($"插件卸载失败: {plugin.Name}");
                        MessageBox.Show($"卸载插件失败: {plugin.Name}", "错误",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    Logger.Info($"用户取消卸载插件: {plugin.Name}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"卸载插件时发生错误: {plugin.Name}", ex);
                MessageBox.Show($"卸载插件时发生错误: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 卸载所有插件
        /// </summary>
        private void UnloadAllPlugins_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Logger.Info("用户请求卸载所有插件");

                var result = MessageBox.Show("确定要卸载所有插件吗？", "确认卸载",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    // 关闭所有标签页
                    WorkspaceTabControl.Items.Clear();

                    // 卸载所有插件
                    _pluginManager.UnloadAllPlugins();

                    UpdateStatus("所有插件已卸载");
                    WorkspaceTitle.Text = "请选择一个插件开始使用";

                    Logger.Info("所有插件卸载完成");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("卸载所有插件时发生错误", ex);
                MessageBox.Show($"卸载所有插件时发生错误: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 关闭指定插件的所有标签页
        /// </summary>
        private void ClosePluginTabs(IPlugin plugin)
        {
            // 查找并关闭该插件的所有标签页
            var tabsToRemove = WorkspaceTabControl.Items
                .Cast<TabItem>()
                .Where(tab => tab.Tag == plugin)
                .ToList();

            foreach (var tab in tabsToRemove)
            {
                WorkspaceTabControl.Items.Remove(tab);
            }

            // 更新工作区标题
            if (WorkspaceTabControl.Items.Count == 0)
            {
                WorkspaceTitle.Text = "请选择一个插件开始使用";
            }
        }

        #endregion

        #region 文件菜单功能

        private void OpenPluginsFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var pluginsRoot = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins");
                if (System.IO.Directory.Exists(pluginsRoot))
                {
                    System.Diagnostics.Process.Start("explorer.exe", pluginsRoot);
                    UpdateStatus("已打开插件根目录");
                }
                else
                {
                    MessageBox.Show("插件目录不存在，将创建目录", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    System.IO.Directory.CreateDirectory(pluginsRoot);
                    System.Diagnostics.Process.Start("explorer.exe", pluginsRoot);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开插件目录失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreatePluginTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "选择插件模板创建位置";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // 使用 WPF 的输入对话框替代 VisualBasic.InputBox
                var inputDialog = new InputDialog("创建插件模板", "请输入插件名称:", "MyNewPlugin");

                if (inputDialog.ShowDialog() == true && !string.IsNullOrEmpty(inputDialog.Answer))
                {
                    var pluginFolder = System.IO.Path.Combine(dialog.SelectedPath, inputDialog.Answer);
                    CreatePluginTemplate(pluginFolder, inputDialog.Answer);
                }
            }
        }

        private void CreatePluginTemplate(string folderPath, string pluginName)
        {
            try
            {
                // 创建文件夹结构
                System.IO.Directory.CreateDirectory(folderPath);
                System.IO.Directory.CreateDirectory(System.IO.Path.Combine(folderPath, "Resources"));
                System.IO.Directory.CreateDirectory(System.IO.Path.Combine(folderPath, "Dependencies"));

                // 创建说明文件
                var readmeContent = $"{pluginName} 插件\n\n" +
                                   "文件夹结构说明:\n" +
                                   $"{pluginName}.dll - 主程序集\n" +
                                   "Resources/ - 资源文件\n" +
                                   "Dependencies/ - 依赖项\n" +
                                   "plugin.json - 配置文件\n\n" +
                                   "开发说明:\n" +
                                   "1. 创建类库项目\n" +
                                   "2. 引用 Toolbox.Core\n" +
                                   "3. 实现 IPlugin 接口\n" +
                                   "4. 编译后将dll复制到此文件夹";

                System.IO.File.WriteAllText(System.IO.Path.Combine(folderPath, "README.txt"), readmeContent, System.Text.Encoding.UTF8);

                // 创建JSON格式的配置文件
                var config = new PluginInfo
                {
                    Name = pluginName,
                    Version = "1.0.0",
                    Description = $"这是一个{pluginName}插件",
                    Author = "YourName",
                    Enabled = true,
                    Dependencies = new List<string>
            {
                "System.Windows.Forms",
                "Newtonsoft.Json"
            },
                    Settings = new Dictionary<string, object>
                    {
                        ["AutoStart"] = false,
                        ["MaxConnections"] = 10,
                        ["Timeout"] = 30
                    }
                };

                // 使用UTF-8编码保存JSON，避免乱码
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(folderPath, "plugin.json"),
                    config.ToJson(),
                    System.Text.Encoding.UTF8
                );

                // 打开创建的文件夹
                System.Diagnostics.Process.Start("explorer.exe", folderPath);

                UpdateStatus($"已创建插件模板: {pluginName}");
                MessageBox.Show($"插件模板已创建在: {folderPath}", "成功",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                Logger.Info($"创建插件模板: {pluginName}");
            }
            catch (Exception ex)
            {
                Logger.Error($"创建插件模板失败: {pluginName}", ex);
                MessageBox.Show($"创建插件模板失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Logger.Info("应用程序关闭");
            // 清理资源
            _pluginManager?.UnloadAllPlugins();
            Application.Current.Shutdown();
        }

        #endregion

        #region 刷新功能
        private void RefreshPluginList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Logger.Info("用户请求刷新插件列表");
                UpdateStatus("刷新插件列表中...");
                _pluginManager.ReloadAllPlugins();
                WorkspaceTabControl.Items.Clear();
                UpdateStatus("插件列表刷新完成");
            }
            catch (Exception ex)
            {
                Logger.Error("刷新插件列表失败", ex);
                MessageBox.Show($"刷新插件列表失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region 日志管理功能

        /// <summary>
        /// 查看日志摘要
        /// </summary>
        private void ViewLogSummary_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var logSummary = Logger.GetRecentLogs(7);
                MessageBox.Show(logSummary, "最近日志摘要",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                Logger.Info("查看日志摘要");
            }
            catch (Exception ex)
            {
                Logger.Error("查看日志摘要失败", ex);
                MessageBox.Show($"查看日志摘要失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 打开日志目录
        /// </summary>
        private void OpenLogsFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var logsRoot = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                if (System.IO.Directory.Exists(logsRoot))
                {
                    System.Diagnostics.Process.Start("explorer.exe", logsRoot);
                    UpdateStatus("已打开日志目录");
                    Logger.Info("打开日志目录");
                }
                else
                {
                    MessageBox.Show("日志目录不存在", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("打开日志目录失败", ex);
                MessageBox.Show($"打开日志目录失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 清理过期日志
        /// </summary>
        private void CleanupOldLogs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var result = MessageBox.Show("确定要清理30天前的日志文件吗？", "确认清理",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    Logger.CleanupOldLogs(30);
                    MessageBox.Show("过期日志清理完成", "完成",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    Logger.Info("清理过期日志");
                    UpdateStatus("过期日志已清理");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("清理过期日志失败", ex);
                MessageBox.Show($"清理过期日志失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region 帮助功能

        private void About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "2317-ToolsBox\n\n" +
                "这是一个支持插件形式的工具箱。\n\n" +
                "Author:\n" +
                "\tWenjieYang",
                "关于工具箱",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        #endregion

        protected override void OnClosed(EventArgs e)
        {
            Logger.Info("应用程序关闭");
            base.OnClosed(e);
        }
    }
}