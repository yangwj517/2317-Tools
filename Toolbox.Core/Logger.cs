using System;
using System.IO;
using System.Text;

namespace Toolbox.Core
{
    /// <summary>
    /// 日志级别
    /// </summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    /// <summary>
    /// 日志管理器
    /// </summary>
    public static class Logger
    {
        private static readonly object _lockObject = new object();
        private static string _logDirectory;

        /// <summary>
        /// 初始化日志系统
        /// </summary>
        public static void Initialize(string logDirectory = "Logs")
        {
            _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logDirectory);

            if (!Directory.Exists(_logDirectory))
            {
                Directory.CreateDirectory(_logDirectory);
            }

            // 记录系统启动日志
            Info("日志系统初始化完成");
        }

        /// <summary>
        /// 记录调试信息
        /// </summary>
        public static void Debug(string message)
        {
            WriteLog(LogLevel.Debug, message);
        }

        /// <summary>
        /// 记录一般信息
        /// </summary>
        public static void Info(string message)
        {
            WriteLog(LogLevel.Info, message);
        }

        /// <summary>
        /// 记录警告信息
        /// </summary>
        public static void Warning(string message)
        {
            WriteLog(LogLevel.Warning, message);
        }

        /// <summary>
        /// 记录错误信息
        /// </summary>
        public static void Error(string message)
        {
            WriteLog(LogLevel.Error, message);
        }

        /// <summary>
        /// 记录错误信息（带异常）
        /// </summary>
        public static void Error(string message, Exception exception)
        {
            var fullMessage = $"{message}{Environment.NewLine}异常信息: {exception.Message}{Environment.NewLine}堆栈跟踪: {exception.StackTrace}";
            WriteLog(LogLevel.Error, fullMessage);
        }

        /// <summary>
        /// 写入日志到文件
        /// </summary>
        private static void WriteLog(LogLevel level, string message)
        {
            lock (_lockObject)
            {
                try
                {
                    // 确保日志目录存在
                    if (string.IsNullOrEmpty(_logDirectory))
                    {
                        Initialize();
                    }

                    // 创建按日期分隔的日志文件
                    string logFile = Path.Combine(_logDirectory, $"log_{DateTime.Now:yyyyMMdd}.txt");

                    // 构建日志条目
                    var logEntry = new StringBuilder();
                    logEntry.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}]");
                    logEntry.AppendLine($"消息: {message}");
                    logEntry.AppendLine(new string('-', 60));

                    // 写入文件
                    File.AppendAllText(logFile, logEntry.ToString(), Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    // 如果日志写入失败，我们无法记录这个错误，只能忽略
                    // 在实际项目中，可能需要回退到其他日志方式
                }
            }
        }

        /// <summary>
        /// 获取最近的日志文件内容
        /// </summary>
        public static string GetRecentLogs(int days = 7)
        {
            try
            {
                var result = new StringBuilder();
                result.AppendLine($"最近 {days} 天的日志摘要");
                result.AppendLine(new string('=', 60));

                for (int i = 0; i < days; i++)
                {
                    var date = DateTime.Now.AddDays(-i);
                    var logFile = Path.Combine(_logDirectory, $"log_{date:yyyyMMdd}.txt");

                    if (File.Exists(logFile))
                    {
                        var fileInfo = new FileInfo(logFile);
                        var logContent = File.ReadAllText(logFile);
                        var lineCount = logContent.Split('\n').Length;

                        result.AppendLine($"{date:yyyy-MM-dd}: {lineCount} 行日志");
                    }
                    else
                    {
                        result.AppendLine($"{date:yyyy-MM-dd}: 无日志");
                    }
                }

                return result.ToString();
            }
            catch (Exception ex)
            {
                return $"读取日志文件时发生错误: {ex.Message}";
            }
        }

        /// <summary>
        /// 清理过期日志文件
        /// </summary>
        public static void CleanupOldLogs(int keepDays = 30)
        {
            try
            {
                if (!Directory.Exists(_logDirectory))
                    return;

                var cutoffDate = DateTime.Now.AddDays(-keepDays);
                var logFiles = Directory.GetFiles(_logDirectory, "log_*.txt");

                foreach (var logFile in logFiles)
                {
                    try
                    {
                        var fileInfo = new FileInfo(logFile);
                        if (fileInfo.CreationTime < cutoffDate)
                        {
                            File.Delete(logFile);
                            Info($"已删除过期日志文件: {Path.GetFileName(logFile)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Error($"删除日志文件失败 {Path.GetFileName(logFile)}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Error($"清理过期日志时发生错误: {ex.Message}");
            }
        }
    }
}