using System.Windows.Controls;
using Toolbox.Core;

namespace ModbusConfigPlugin
{
    public class ModbusConfigPlugin : IPlugin
    {
        public string Name => "ELC-ModbusTcp配置生成器";
        public string Version => "1.0.0";
        public string Description => "从Excel生成Modbus TCP通信配置文件";
        public string Author => "WenjieYang";
        public string ToolName => "ModbusConfigPlugin";

        public UserControl GetControl()
        {
            return new ModbusConfigControl();
        }

        public void Initialize()
        {
            // 初始化逻辑
            Logger.Info("加载 ELC-Modbus-Tcp 配置生成插件");
        }

        public void Dispose()
        {
            // 清理逻辑
            Logger.Info("ELC-Modbus-Tcp 配置生成插件已卸载");
        }

        internal class pojo
        {
        }
    }
}