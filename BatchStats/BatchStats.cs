using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Toolbox.Core;

namespace BatchStats
{
    public class BatchStats : IPlugin

    {
        public string Name => "Stats文件批量导出表格";
        public string Version => "1.0.0";
        public string Description => "从Stats文件导出Excel文件";
        public string Author => "WenjieYang";
        public string ToolName => "BatchStats";

        public UserControl GetControl()
        {
            return new BatchStatsControl();
        }

        public void Initialize()
        {
            // 初始化逻辑
            Logger.Info("加载 BatchStats 插件");
        }

        public void Dispose()
        {
            // 清理逻辑
            Logger.Info("BatchStats 插件已卸载");
        }
    }
}
