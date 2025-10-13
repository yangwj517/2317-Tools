using System.Windows.Controls;
using Toolbox.Core;

namespace CalculatorPlugin
{
    public class CalculatorPlugin : IPlugin
    {
        public string Name => "计算器";
        public string Version => "1.0.0";
        public string Description => "一个简单的计算器插件，支持基本数学运算";
        public string Author => "工具箱开发者";

        public string ToolName => "CalculatorPlugin";

        public UserControl GetControl()
        {
            Logger.Info("计算器插件: 创建用户界面");
            return new CalculatorControl();
        }

        public void Initialize()
        {
            // 可以在这里进行插件初始化，比如加载配置等
            Logger.Info("计算器插件已初始化");
        }

        public void Dispose()
        {
            // 清理资源
            Logger.Info("计算器插件已清理");
        }
    }
}