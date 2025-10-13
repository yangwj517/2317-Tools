using System.Windows.Controls;

namespace Toolbox.Core
{
    /// <summary>
    /// 插件接口 - 所有插件都必须实现这个接口
    /// </summary>
    public interface IPlugin
    {
        /// <summary>
        /// 插件名称（显示在列表中）
        /// </summary>
        string Name { get; }

        /// <summary>
        /// 插件版本
        /// </summary>
        string Version { get; }

        /// <summary>
        /// 插件描述
        /// </summary>
        string Description { get; }

        /// <summary>
        /// 插件作者
        /// </summary>
        string Author { get; }

        /// <summary>
        /// 工具名称 == 》 与程序包名一致
        /// </summary>
        string ToolName { get; }

        /// <summary>
        /// 获取插件的用户界面
        /// </summary>
        /// <returns>WPF用户控件</returns>
        UserControl GetControl();

        /// <summary>
        /// 初始化插件
        /// </summary>
        void Initialize();

        /// <summary>
        /// 清理插件资源
        /// </summary>
        void Dispose();
    }
}