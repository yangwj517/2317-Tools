using System;
using System.IO;
using System.Windows;

namespace ToolboxHost
{
    /// <summary>
    /// 插件文件夹辅助工具
    /// </summary>
    public static class PluginFolderHelper
    {
        /// <summary>
        /// 为插件创建标准文件夹结构
        /// </summary>
        /// <param name="pluginName">插件名称</param>
        /// <returns>是否创建成功</returns>
        public static bool CreatePluginFolderStructure(string pluginName)
        {
            try
            {
                var pluginsRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins");
                var pluginFolder = Path.Combine(pluginsRoot, pluginName);

                // 创建主文件夹
                Directory.CreateDirectory(pluginFolder);

                // 创建子文件夹
                Directory.CreateDirectory(Path.Combine(pluginFolder, "Resources"));
                Directory.CreateDirectory(Path.Combine(pluginFolder, "Resources", "Images"));
                Directory.CreateDirectory(Path.Combine(pluginFolder, "Resources", "Data"));
                Directory.CreateDirectory(Path.Combine(pluginFolder, "Dependencies"));

                // 创建配置文件模板
                var configContent = $"# {pluginName} 插件配置\n" +
                                   $"Version=1.0\n" +
                                   $"Author=YourName\n" +
                                   $"Description=插件描述\n";

                File.WriteAllText(Path.Combine(pluginFolder, "config.txt"), configContent);

                // 创建说明文件
                var readmeContent = $"插件文件夹结构:\n" +
                                   $"{pluginName}.dll - 主程序集 (放在此文件夹根目录)\n" +
                                   $"Resources/ - 资源文件目录\n" +
                                   $"Dependencies/ - 依赖项目录\n" +
                                   $"config.txt - 配置文件\n";

                File.WriteAllText(Path.Combine(pluginFolder, "README.txt"), readmeContent);

                Console.WriteLine($"已创建插件文件夹结构: {pluginFolder}");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建插件文件夹失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// 打开插件文件夹
        /// </summary>
        /// <param name="pluginName">插件名称</param>
        public static void OpenPluginFolder(string pluginName)
        {
            try
            {
                var pluginFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins", pluginName);
                if (Directory.Exists(pluginFolder))
                {
                    System.Diagnostics.Process.Start("explorer.exe", pluginFolder);
                }
                else
                {
                    MessageBox.Show($"插件文件夹不存在: {pluginFolder}", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开插件文件夹失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 获取所有插件文件夹
        /// </summary>
        /// <returns>插件文件夹列表</returns>
        public static string[] GetPluginFolders()
        {
            var pluginsRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins");
            if (Directory.Exists(pluginsRoot))
            {
                return Directory.GetDirectories(pluginsRoot);
            }
            return new string[0];
        }
    }
}