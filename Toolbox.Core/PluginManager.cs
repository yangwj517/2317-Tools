using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;

namespace Toolbox.Core
{
    public class PluginManager
    {
        private List<IPlugin> _plugins;
        private Dictionary<IPlugin, PluginInfo> _pluginConfigs;
        private Dictionary<IPlugin, PluginLoadContext> _pluginLoadContexts;
        private string _pluginsRootDirectory;
        // 新增：跟踪所有临时文件路径，确保能彻底清理
        private HashSet<string> _allShadowCopiedPaths = new HashSet<string>();

        public class PluginLoadContext
        {
            public string FolderPath { get; set; }
            public string ConfigFilePath { get; set; }
            public string ResourcesPath { get; set; }
            public Assembly Assembly { get; set; }
            public DateTime LoadTime { get; set; }
            public string DependenciesPath { get; set; }
            public string ShadowCopiedDllPath { get; set; }
        }

        // 保留原有属性和事件
        public IReadOnlyList<IPlugin> Plugins => _plugins.AsReadOnly();
        public IReadOnlyDictionary<IPlugin, PluginInfo> PluginConfigs => _pluginConfigs;
        public IReadOnlyDictionary<IPlugin, PluginLoadContext> PluginLoadContexts => _pluginLoadContexts;
        public event Action<IPlugin, PluginInfo> PluginLoaded;
        public event Action<IPlugin, PluginInfo> PluginUnloaded;

        public PluginManager(string pluginsRootDirectory = "Plugins")
        {
            _plugins = new List<IPlugin>();
            _pluginConfigs = new Dictionary<IPlugin, PluginInfo>();
            _pluginLoadContexts = new Dictionary<IPlugin, PluginLoadContext>();
            _pluginsRootDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, pluginsRootDirectory);

            EnsurePluginDirectoryStructure();
            // 启动时清理可能残留的临时文件
            CleanupAllShadowDirectories();

            Logger.Info($"插件管理器初始化完成，插件目录: {_pluginsRootDirectory}");
        }

        private void EnsurePluginDirectoryStructure()
        {
            if (!Directory.Exists(_pluginsRootDirectory))
            {
                Directory.CreateDirectory(_pluginsRootDirectory);
                Logger.Info($"创建插件根目录: {_pluginsRootDirectory}");
            }
        }

        public void LoadAllPlugins()
        {
            Logger.Info($"开始扫描插件目录: {_pluginsRootDirectory}");

            if (!Directory.Exists(_pluginsRootDirectory))
            {
                Logger.Warning("插件根目录不存在，跳过加载");
                return;
            }

            var pluginFolders = Directory.GetDirectories(_pluginsRootDirectory)
                .Where(folder => !new DirectoryInfo(folder).Attributes.HasFlag(FileAttributes.Hidden))
                .ToList();

            Logger.Info($"找到 {pluginFolders.Count} 个插件文件夹");

            foreach (var pluginFolder in pluginFolders)
            {
                try
                {
                    LoadPluginFromFolder(pluginFolder);
                }
                catch (Exception ex)
                {
                    Logger.Error($"加载插件失败 {Path.GetFileName(pluginFolder)}", ex);
                }
            }

            Logger.Info($"插件加载完成，共加载 {_plugins.Count} 个插件");
        }

        private void LoadPluginFromFolder(string pluginFolder)
        {
            var folderName = Path.GetFileName(pluginFolder);
            Logger.Info($"扫描插件文件夹: {folderName}");

            SetupPluginDependencyResolution(pluginFolder);
            var pluginConfig = LoadPluginConfig(pluginFolder);

            if (!pluginConfig.Enabled)
            {
                Logger.Info($"插件 {folderName} 已被禁用，跳过加载");
                return;
            }

            var expectedDllName = $"{folderName}.dll";
            var pluginDllPath = Path.Combine(pluginFolder, expectedDllName);

            if (!File.Exists(pluginDllPath))
            {
                var dllFiles = Directory.GetFiles(pluginFolder, "*.dll");
                if (dllFiles.Length == 0)
                {
                    Logger.Warning($"在文件夹 {folderName} 中未找到DLL文件");
                    return;
                }
                pluginDllPath = dllFiles[0];
                Logger.Info($"使用找到的DLL: {Path.GetFileName(pluginDllPath)}");
            }

            // 关键修复1：使用唯一GUID生成临时文件名，避免冲突
            string tempShadowDir = Path.Combine(Path.GetTempPath(), "ToolboxPluginShadow",
                $"{folderName}_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempShadowDir);

            string shadowCopiedDllPath = Path.Combine(tempShadowDir, Path.GetFileName(pluginDllPath));

            try
            {
                // 关键修复2：多次尝试复制文件，解决临时锁定问题
                bool copySuccess = false;
                int retryCount = 0;
                while (!copySuccess && retryCount < 5)
                {
                    try
                    {
                        // 先删除可能存在的旧文件
                        if (File.Exists(shadowCopiedDllPath))
                        {
                            File.Delete(shadowCopiedDllPath);
                        }
                        File.Copy(pluginDllPath, shadowCopiedDllPath, true);
                        copySuccess = true;
                    }
                    catch (IOException)
                    {
                        // 等待100ms后重试
                        retryCount++;
                        System.Threading.Thread.Sleep(100);
                    }
                }

                if (!copySuccess)
                {
                    Logger.Error($"多次尝试后仍无法复制文件: {pluginDllPath}");
                    return;
                }

                // 记录临时文件路径，用于后续清理
                lock (_allShadowCopiedPaths)
                {
                    _allShadowCopiedPaths.Add(shadowCopiedDllPath);
                }

                var loadContext = new PluginLoadContext
                {
                    FolderPath = pluginFolder,
                    ConfigFilePath = Path.Combine(pluginFolder, "plugin.json"),
                    ResourcesPath = Path.Combine(pluginFolder, "Resources"),
                    DependenciesPath = Path.Combine(pluginFolder, "Dependencies"),
                    LoadTime = DateTime.Now,
                    ShadowCopiedDllPath = shadowCopiedDllPath
                };

                var assemblyLoader = new PluginAssemblyLoader(pluginFolder);
                var assembly = assemblyLoader.LoadFromAssemblyPath(shadowCopiedDllPath);
                loadContext.Assembly = assembly;

                bool pluginFound = false;
                foreach (Type type in assembly.GetTypes())
                {
                    if (typeof(IPlugin).IsAssignableFrom(type) && !type.IsInterface && !type.IsAbstract)
                    {
                        Logger.Info($"找到插件类型: {type.FullName}");

                        IPlugin plugin = (IPlugin)Activator.CreateInstance(type);
                        plugin.Initialize();

                        _plugins.Add(plugin);
                        _pluginConfigs[plugin] = pluginConfig;
                        _pluginLoadContexts[plugin] = loadContext;
                        PluginLoaded?.Invoke(plugin, pluginConfig);

                        Logger.Info($"插件加载成功: {plugin.Name} v{plugin.Version}");
                        pluginFound = true;
                        break;
                    }
                }

                if (!pluginFound)
                {
                    Logger.Warning($"在 {Path.GetFileName(pluginDllPath)} 中未找到实现IPlugin接口的类型");
                    DeleteShadowCopyFiles(loadContext);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"加载插件失败 {folderName}", ex);
                DeleteShadowCopyFiles(new PluginLoadContext { ShadowCopiedDllPath = shadowCopiedDllPath });
            }
        }

        public class PluginAssemblyLoader
        {
            private readonly string _pluginFolder;
            private readonly string _dependenciesPath;

            public PluginAssemblyLoader(string pluginFolder)
            {
                _pluginFolder = pluginFolder;
                _dependenciesPath = Path.Combine(pluginFolder, "Dependencies");
            }

            public Assembly LoadFromAssemblyPath(string assemblyPath)
            {
                AppDomain.CurrentDomain.AssemblyResolve += ResolveAssembly;

                try
                {
                    return Assembly.LoadFrom(assemblyPath);
                }
                finally
                {
                    AppDomain.CurrentDomain.AssemblyResolve -= ResolveAssembly;
                }
            }

            private Assembly ResolveAssembly(object sender, ResolveEventArgs args)
            {
                try
                {
                    var assemblyName = new AssemblyName(args.Name);
                    var assemblyPath = Path.Combine(_pluginFolder, assemblyName.Name + ".dll");
                    if (File.Exists(assemblyPath))
                    {
                        return Assembly.LoadFrom(assemblyPath);
                    }

                    if (Directory.Exists(_dependenciesPath))
                    {
                        assemblyPath = Path.Combine(_dependenciesPath, assemblyName.Name + ".dll");
                        if (File.Exists(assemblyPath))
                        {
                            return Assembly.LoadFrom(assemblyPath);
                        }

                        var dllFiles = Directory.GetFiles(_dependenciesPath, "*.dll");
                        foreach (var dllFile in dllFiles)
                        {
                            var fileName = Path.GetFileNameWithoutExtension(dllFile);
                            if (fileName.StartsWith(assemblyName.Name))
                            {
                                return Assembly.LoadFrom(dllFile);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"解析程序集失败: {args.Name}", ex);
                }

                return null;
            }
        }

        private void SetupPluginDependencyResolution(string pluginFolder)
        {
            var dependenciesPath = Path.Combine(pluginFolder, "Dependencies");

            if (Directory.Exists(dependenciesPath))
            {
                AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
                {
                    try
                    {
                        var assemblyName = new AssemblyName(args.Name).Name + ".dll";
                        var assemblyPath = Path.Combine(dependenciesPath, assemblyName);

                        if (File.Exists(assemblyPath))
                        {
                            Logger.Info($"从依赖目录加载: {assemblyName}");
                            return Assembly.LoadFrom(assemblyPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"解析程序集依赖失败: {args.Name}", ex);
                    }
                    return null;
                };
            }
        }

        private PluginInfo LoadPluginConfig(string pluginFolder)
        {
            var configPath = Path.Combine(pluginFolder, "plugin.json");

            if (File.Exists(configPath))
            {
                try
                {
                    var json = File.ReadAllText(configPath, System.Text.Encoding.UTF8);
                    var config = PluginInfo.FromJson(json);
                    Logger.Info($"已加载插件配置: {Path.GetFileName(pluginFolder)}");

                    if (config.Dependencies == null)
                    {
                        config.Dependencies = new List<string>();
                    }

                    return config;
                }
                catch (Exception ex)
                {
                    Logger.Error($"加载插件配置失败 {Path.GetFileName(pluginFolder)}", ex);
                }
            }
            else
            {
                Logger.Info($"插件配置不存在，创建默认配置: {Path.GetFileName(pluginFolder)}");
                var defaultConfig = new PluginInfo
                {
                    Name = Path.GetFileName(pluginFolder),
                    Version = "1.0.0",
                    Description = "插件描述",
                    Author = "未知作者",
                    Enabled = true,
                    Dependencies = new List<string>(),
                    Settings = new Dictionary<string, object>()
                };

                SavePluginConfig(pluginFolder, defaultConfig);
                return defaultConfig;
            }

            return new PluginInfo
            {
                Name = Path.GetFileName(pluginFolder),
                Version = "1.0.0",
                Description = "插件描述",
                Author = "未知作者",
                Enabled = true,
                Dependencies = new List<string>(),
                Settings = new Dictionary<string, object>()
            };
        }

        public bool SavePluginConfig(IPlugin plugin, PluginInfo config)
        {
            if (_pluginLoadContexts.TryGetValue(plugin, out var loadContext))
            {
                return SavePluginConfig(loadContext.FolderPath, config);
            }
            return false;
        }

        private bool SavePluginConfig(string pluginFolder, PluginInfo config)
        {
            try
            {
                var configPath = Path.Combine(pluginFolder, "plugin.json");
                var json = config.ToJson();
                File.WriteAllText(configPath, json, System.Text.Encoding.UTF8);
                Logger.Info($"插件配置已保存: {Path.GetFileName(pluginFolder)}");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"保存插件配置失败 {Path.GetFileName(pluginFolder)}", ex);
                return false;
            }
        }

        public PluginInfo GetPluginConfig(IPlugin plugin)
        {
            if (_pluginConfigs.TryGetValue(plugin, out var config))
            {
                return config;
            }
            return new PluginInfo();
        }

        public PluginLoadContext GetPluginLoadContext(IPlugin plugin)
        {
            if (_pluginLoadContexts.TryGetValue(plugin, out var loadContext))
            {
                return loadContext;
            }
            return null;
        }

        public bool UpdatePluginConfig(IPlugin plugin, Action<PluginInfo> updateAction)
        {
            if (_pluginConfigs.TryGetValue(plugin, out var config))
            {
                updateAction(config);
                config.Updated = DateTime.Now;
                return SavePluginConfig(plugin, config);
            }
            return false;
        }

        public bool UnloadPlugin(IPlugin plugin)
        {
            try
            {
                if (_plugins.Contains(plugin) && _pluginConfigs.TryGetValue(plugin, out var pluginConfig))
                {
                    Logger.Info($"开始卸载插件: {plugin.Name}");

                    // 1. 调用插件清理方法
                    plugin.Dispose();

                    // 2. 新增：释放插件依赖文件的锁定
                    if (_pluginLoadContexts.TryGetValue(plugin, out var loadContext) &&
                        Directory.Exists(loadContext.DependenciesPath))
                    {
                        ReleaseDirectoryLocks(loadContext.DependenciesPath);
                    }

                    // 3. 清理临时文件
                    if (_pluginLoadContexts.TryGetValue(plugin, out var context))
                    {
                        DeleteShadowCopyFiles(context);
                    }

                    // 4. 移除引用
                    _plugins.Remove(plugin);
                    _pluginConfigs.Remove(plugin);
                    _pluginLoadContexts.Remove(plugin);

                    // 5. 触发事件
                    PluginUnloaded?.Invoke(plugin, pluginConfig);

                    Logger.Info($"插件卸载成功: {plugin.Name}");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"卸载插件失败 {plugin.Name}", ex);
                return false;
            }
        }

        /// <summary>
        /// 释放目录中所有文件的锁定
        /// </summary>
        private void ReleaseDirectoryLocks(string directoryPath)
        {
            try
            {
                foreach (var file in Directory.GetFiles(directoryPath))
                {
                    ReleaseFileLock(file);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"释放目录锁定时出错: {directoryPath}, 错误: {ex.Message}");
            }
        }

        // <summary>
        /// 尝试释放文件锁定
        /// </summary>
        private void ReleaseFileLock(string filePath)
        {
            try
            {
                // 检查文件是否被锁定
                if (IsFileLocked(filePath))
                {
                    Logger.Info($"尝试释放文件锁定: {filePath}");

                    // 方法1: 尝试以只读方式打开并立即关闭
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Delete))
                    {
                        stream.Close();
                    }

                    // 方法2: 强制垃圾回收，可能释放文件句柄
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch
            {
                // 释放失败时不抛出异常，继续重试
            }

        }

        /// <summary>
        /// 检查文件是否被锁定
        /// </summary>
        private bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                using (File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException ex)
            {
                return IsFileLocked(ex);
            }
        }

        /// <summary>
        /// 检查文件是否被锁定
        /// </summary>
        private bool IsFileLocked(Exception ex)
        {
            int errorCode = System.Runtime.InteropServices.Marshal.GetHRForException(ex) & ((1 << 16) - 1);
            return errorCode == 32 || errorCode == 33;
        }


        public bool UnloadPluginByName(string pluginName)
        {
            var plugin = _plugins.FirstOrDefault(p => p.Name.Equals(pluginName, StringComparison.OrdinalIgnoreCase));
            if (plugin != null)
            {
                return UnloadPlugin(plugin);
            }

            Logger.Warning($"未找到名为 '{pluginName}' 的插件");
            return false;
        }

        public void UnloadAllPlugins()
        {
            Logger.Info("开始卸载所有插件...");

            var pluginsToUnload = _plugins.ToList();
            foreach (var plugin in pluginsToUnload)
            {
                try
                {
                    if (_pluginConfigs.TryGetValue(plugin, out var config))
                    {
                        plugin.Dispose();
                        if (_pluginLoadContexts.TryGetValue(plugin, out var loadContext))
                        {
                            DeleteShadowCopyFiles(loadContext);
                        }
                        PluginUnloaded?.Invoke(plugin, config);
                        _pluginConfigs.Remove(plugin);
                        _pluginLoadContexts.Remove(plugin);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"卸载插件失败 {plugin.Name}", ex);
                }
            }

            _plugins.Clear();
            Logger.Info("所有插件已卸载");
        }

        // 关键修复3：增强重新加载逻辑
        public void ReloadAllPlugins()
        {
            Logger.Info("开始重新加载所有插件");

            // 1. 卸载所有插件
            UnloadAllPlugins();

            // 2. 强制清理所有临时文件（带重试机制）
            CleanupAllShadowDirectories(true);

            // 3. 清理程序集缓存
            ClearAssemblyCache();

            // 4. 重新加载插件
            LoadAllPlugins();

            Logger.Info("插件重新加载完成");
        }

        public string GetPluginResourcePath(IPlugin plugin, string relativePath)
        {
            if (_pluginLoadContexts.TryGetValue(plugin, out var loadContext) &&
                Directory.Exists(loadContext.ResourcesPath))
            {
                return Path.Combine(loadContext.ResourcesPath, relativePath);
            }
            return null;
        }

        public string[] GetPluginFiles(IPlugin plugin, string searchPattern = "*.*")
        {
            if (_pluginLoadContexts.TryGetValue(plugin, out var loadContext))
            {
                return Directory.GetFiles(loadContext.FolderPath, searchPattern, SearchOption.AllDirectories);
            }
            return new string[0];
        }

        // 关键修复4：增强临时文件删除逻辑，带重试机制
        private void DeleteShadowCopyFiles(PluginLoadContext loadContext)
        {
            if (string.IsNullOrEmpty(loadContext.ShadowCopiedDllPath))
                return;

            try
            {
                // 从跟踪集合中移除
                lock (_allShadowCopiedPaths)
                {
                    _allShadowCopiedPaths.Remove(loadContext.ShadowCopiedDllPath);
                }

                // 多次尝试删除文件
                int retryCount = 0;
                bool deleted = false;
                while (!deleted && retryCount < 5)
                {
                    try
                    {
                        if (File.Exists(loadContext.ShadowCopiedDllPath))
                        {
                            File.Delete(loadContext.ShadowCopiedDllPath);
                        }
                        deleted = true;
                    }
                    catch (IOException)
                    {
                        retryCount++;
                        System.Threading.Thread.Sleep(100);
                    }
                }

                if (!deleted)
                {
                    Logger.Warning($"无法删除临时文件: {loadContext.ShadowCopiedDllPath}");
                    return;
                }

                // 删除临时目录
                string tempDir = Path.GetDirectoryName(loadContext.ShadowCopiedDllPath);
                if (Directory.Exists(tempDir))
                {
                    // 先尝试正常删除
                    try
                    {
                        Directory.Delete(tempDir, true);
                    }
                    catch (IOException)
                    {
                        // 延迟后再试一次
                        System.Threading.Thread.Sleep(200);
                        try
                        {
                            Directory.Delete(tempDir, true);
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"无法删除临时目录: {tempDir}, 原因: {ex.Message}");
                        }
                    }
                }

                Logger.Info($"已清理插件临时文件: {loadContext.ShadowCopiedDllPath}");
            }
            catch (Exception ex)
            {
                Logger.Warning($"清理临时文件失败: {ex.Message}");
            }
        }

        // 关键修复5：增强全局临时目录清理
        private void CleanupAllShadowDirectories(bool force = false)
        {
            try
            {
                string baseShadowDir = Path.Combine(Path.GetTempPath(), "ToolboxPluginShadow");
                if (!Directory.Exists(baseShadowDir))
                    return;

                // 先处理跟踪的文件
                lock (_allShadowCopiedPaths)
                {
                    var pathsToDelete = _allShadowCopiedPaths.ToList();
                    foreach (var path in pathsToDelete)
                    {
                        DeleteShadowCopyFiles(new PluginLoadContext { ShadowCopiedDllPath = path });
                    }
                    _allShadowCopiedPaths.Clear();
                }

                // 强制删除整个目录（如果需要）
                if (force && Directory.Exists(baseShadowDir))
                {
                    int retryCount = 0;
                    bool deleted = false;
                    while (!deleted && retryCount < 5)
                    {
                        try
                        {
                            Directory.Delete(baseShadowDir, true);
                            deleted = true;
                        }
                        catch (IOException)
                        {
                            retryCount++;
                            System.Threading.Thread.Sleep(200);
                        }
                    }

                    if (deleted)
                    {
                        Logger.Info("已强制清理所有插件的影子复制临时目录");
                    }
                    else
                    {
                        Logger.Warning("无法完全清理影子复制目录，部分文件可能仍被锁定");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"清理影子复制目录失败: {ex.Message}");
            }
        }

        private void ClearAssemblyCache()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                Logger.Info("已清理程序集缓存");
            }
            catch (Exception ex)
            {
                Logger.Warning($"清理程序集缓存失败: {ex.Message}");
            }
        }
    }
}
