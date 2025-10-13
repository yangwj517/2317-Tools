using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Toolbox.Core;

namespace ToolboxHost
{
    public partial class PluginInfoWindow : Window
    {
        public string PluginName { get; }
        public string PluginVersion { get; }
        public string PluginDescription { get; }
        public string PluginAuthor { get; }
        public string PluginEnabled { get; }
        public string PluginCreated { get; }
        public string PluginUpdated { get; }
        public List<string> PluginDependencies { get; }
        public List<KeyValueDisplay> PluginSettings { get; }

        public PluginInfoWindow(IPlugin plugin, PluginInfo config)
        {
            // 直接从配置初始化所有属性
            PluginName = config.Name;
            PluginVersion = config.Version;
            PluginDescription = config.Description;
            PluginAuthor = config.Author;
            PluginEnabled = config.Enabled ? "已启用" : "已禁用";
            PluginCreated = config.Created.ToString("yyyy-MM-dd HH:mm:ss");
            PluginUpdated = config.Updated.ToString("yyyy-MM-dd HH:mm:ss");
            PluginDependencies = config.Dependencies ?? new List<string>();
            PluginSettings = config.Settings?.Select(kv => new KeyValueDisplay
            {
                Key = kv.Key,
                Value = kv.Value?.ToString() ?? "null"
            }).ToList() ?? new List<KeyValueDisplay>();

            InitializeComponent();
            DataContext = this;

            Title = $"插件信息 - {PluginName}";

            // 更新可见性
            UpdateVisibility();
        }

        private void UpdateVisibility()
        {
            // 更新依赖项可见性
            if (PluginDependencies.Count > 0)
            {
                NoDependenciesText.Visibility = Visibility.Collapsed;
            }
            else
            {
                NoDependenciesText.Visibility = Visibility.Visible;
            }

            // 更新设置可见性
            if (PluginSettings.Count > 0)
            {
                NoSettingsText.Visibility = Visibility.Collapsed;
            }
            else
            {
                NoSettingsText.Visibility = Visibility.Visible;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class KeyValueDisplay
    {
        public string Key { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }
}