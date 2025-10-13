using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace BatchStats
{
    /// <summary>
    /// 日志类型转颜色转换器（XAML绑定专用）
    /// </summary>
    public class LogTypeToColorConverter : IValueConverter
    {
        /// <summary>
        /// 正向转换：日志文本 → 对应颜色
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // 安全转换为string
            string logText = value as string;
            if (string.IsNullOrEmpty(logText))
                return Brushes.Black;

            // 根据日志类型返回颜色
            if (logText.Contains("[ERROR]"))
                return Brushes.Red;
            else if (logText.Contains("[WARNING]"))
                return Brushes.Orange;
            else if (logText.Contains("[SUCCESS]"))
                return Brushes.Green;
            else
                return Brushes.Black;
        }

        /// <summary>
        /// 反向转换（无需实现）
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException("无需反向转换");
        }
    }
}
