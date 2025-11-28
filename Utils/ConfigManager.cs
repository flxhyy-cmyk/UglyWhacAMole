using System;
using System.IO;
using Newtonsoft.Json;
using WindowInspector.Models;

namespace WindowInspector.Utils
{
    public class ConfigManager
    {
        private readonly string _programDir;
        private readonly string _configsDir;
        private readonly string _lastConfigFile;
        private readonly string _windowPositionFile;

        public string ConfigsDirectory => _configsDir;
        public string ProgramDirectory => _programDir;

        public ConfigManager()
        {
            // 使用系统的用户应用数据文件夹
            var appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _programDir = Path.Combine(appDataDir, "WindowInspector");
            _configsDir = Path.Combine(_programDir, "configs");
            _lastConfigFile = Path.Combine(_programDir, "last_config.json");
            _windowPositionFile = Path.Combine(_programDir, "window_position.json");

            // 确保目录存在
            if (!Directory.Exists(_programDir))
                Directory.CreateDirectory(_programDir);
            if (!Directory.Exists(_configsDir))
                Directory.CreateDirectory(_configsDir);
        }

        public void SaveConfig(WindowConfig config, string? configName = null)
        {
            try
            {
                var fileName = configName ?? "window_config.json";
                var filePath = Path.Combine(_programDir, fileName);
                var json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"保存配置失败: {ex.Message}");
            }
        }

        public WindowConfig? LoadConfig(string? configName = null)
        {
            try
            {
                var fileName = configName ?? "window_config.json";
                var filePath = Path.Combine(_programDir, fileName);
                
                if (!File.Exists(filePath))
                    return null;

                var json = File.ReadAllText(filePath);
                return JsonConvert.DeserializeObject<WindowConfig>(json);
            }
            catch
            {
                return null;
            }
        }

        public void SaveWindowPosition(WindowPosition position)
        {
            try
            {
                var json = JsonConvert.SerializeObject(position, Formatting.Indented);
                File.WriteAllText(_windowPositionFile, json);
            }
            catch { }
        }

        public WindowPosition? LoadWindowPosition()
        {
            try
            {
                if (!File.Exists(_windowPositionFile))
                    return null;

                var json = File.ReadAllText(_windowPositionFile);
                return JsonConvert.DeserializeObject<WindowPosition>(json);
            }
            catch
            {
                return null;
            }
        }

        public void SaveLastConfig(string configName)
        {
            try
            {
                File.WriteAllText(_lastConfigFile, configName);
            }
            catch { }
        }

        public string? LoadLastConfig()
        {
            try
            {
                if (!File.Exists(_lastConfigFile))
                    return null;
                return File.ReadAllText(_lastConfigFile);
            }
            catch
            {
                return null;
            }
        }
    }
}
