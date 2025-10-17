using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Toolbox.Core;

namespace TempBatch
{
    public class TempBatch : IPlugin
    {

        public string Name => "模板批量创建";
        public string Version => "1.0.0";
        public string Description => "根据模板文件批量生成";
        public string Author => "WenjieYang";
        public string ToolName => "TempBatch";

        public UserControl GetControl()
        {
            return new TempBatchControl();
        }

        public void Initialize()
        {
            // 初始化逻辑
            Logger.Info("加载文件批量创建插件");
        }

        public void Dispose()
        {
            // 清理逻辑
            Logger.Info("文件批量创建插件已卸载");
        }

    }
}
