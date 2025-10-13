using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Toolbox.Core
{
    /// <summary>
    /// 插件信息类 - 用于JSON序列化
    /// </summary>
    public class PluginInfo
    {
        [JsonProperty("name")]
        public string Name { get; set; } = string.Empty;

        [JsonProperty("version")]
        public string Version { get; set; } = string.Empty;

        [JsonProperty("description")]
        public string Description { get; set; } = string.Empty;

        [JsonProperty("author")]
        public string Author { get; set; } = string.Empty;

        [JsonProperty("enabled")]
        public bool Enabled { get; set; } = true;

        [JsonProperty("dependencies")]
        public List<string> Dependencies { get; set; } = new List<string>();

        [JsonProperty("settings")]
        public Dictionary<string, object> Settings { get; set; } = new Dictionary<string, object>();

        [JsonProperty("created")]
        public DateTime Created { get; set; } = DateTime.Now;

        [JsonProperty("updated")]
        public DateTime Updated { get; set; } = DateTime.Now;

        /// <summary>
        /// 从JSON字符串创建PluginInfo
        /// </summary>
        public static PluginInfo FromJson(string json)
        {
            try
            {
                Logger.Info($"开始解析JSON: {json}");

                if (string.IsNullOrWhiteSpace(json))
                {
                    Logger.Warning("JSON字符串为空");
                    return new PluginInfo();
                }

                var settings = new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    DateParseHandling = DateParseHandling.None
                };

                var pluginInfo = JsonConvert.DeserializeObject<PluginInfo>(json, settings);

                if (pluginInfo == null)
                {
                    Logger.Warning("JSON反序列化返回null");
                    return new PluginInfo();
                }

                // 确保列表和字典不为null
                pluginInfo.Dependencies = pluginInfo.Dependencies ?? new List<string>();
                pluginInfo.Settings = pluginInfo.Settings ?? new Dictionary<string, object>();

                Logger.Info($"JSON解析成功 - 名称: '{pluginInfo.Name}', 版本: '{pluginInfo.Version}'");
                Logger.Info($"依赖项数量: {pluginInfo.Dependencies.Count}, 设置数量: {pluginInfo.Settings.Count}");

                return pluginInfo;
            }
            catch (Exception ex)
            {
                Logger.Error($"解析插件信息JSON失败: {ex.Message}");
                Logger.Error($"异常详情: {ex}");
                return new PluginInfo();
            }
        }

        /// <summary>
        /// 转换为JSON字符串
        /// </summary>
        public string ToJson()
        {
            try
            {
                var settings = new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore,
                    DateFormatString = "yyyy-MM-ddTHH:mm:ss"
                };

                var json = JsonConvert.SerializeObject(this, settings);
                Logger.Info($"JSON序列化成功，长度: {json.Length}");
                return json;
            }
            catch (Exception ex)
            {
                Logger.Error($"序列化插件信息为JSON失败: {ex.Message}");
                return "{}";
            }
        }

        /// <summary>
        /// 从IPlugin接口创建PluginInfo
        /// </summary>
        public static PluginInfo FromPlugin(IPlugin plugin)
        {
            return new PluginInfo
            {
                Name = plugin.Name,
                Version = plugin.Version,
                Description = plugin.Description,
                Author = plugin.Author,
                Updated = DateTime.Now
            };
        }
    }
}