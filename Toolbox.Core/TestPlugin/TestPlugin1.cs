using System;
using System.Windows;
using System.Windows.Controls;
using Toolbox.Core;

namespace Toolbox.Core.TestPlugin
{
    public class TestPlugin1 : IPlugin
    {
        public string Name => "测试插件1";
        
        public string Version => "1.0.0";
        
        public string Description => "这是一个测试插件";
        
        public string Author => "开发者";
        
        public string ToolName => "TestPlugin1";

        public UserControl GetControl()
        {
            var control = new UserControl();
            control.Content = new System.Windows.Controls.Label
            {
                Content = "这是测试插件1的界面",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            return control;
        }

        public void Initialize()
        {
            // 插件初始化逻辑
            Console.WriteLine("测试插件1已初始化");
        }

        public void Dispose()
        {
            // 插件清理逻辑
            Console.WriteLine("测试插件1已清理");
        }
    }
}