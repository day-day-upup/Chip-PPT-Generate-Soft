using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.DataSourceVersioning;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.IO;
namespace ChipManualGenerationSogt
{

    public class Settings
    { 
        public string LastLoggedInUser { get; set; }

        public string LastLoggedInPassword { get; set; }



    }

    public class JsonPaser
    {

        public static async Task CreateSettingsJsonFile(string filePath, Settings sessionData)
        {
            // 1. 实例化对象并填充数据
           

            // 2. 配置序列化选项 (可选, 只是为了美观输出)
            var options = new JsonSerializerOptions
            {
                // 使 JSON 输出格式化 (带缩进和换行)
                WriteIndented = true
            };

            try
            {
                // 3. 序列化对象为 JSON 字符串
                string jsonString = JsonSerializer.Serialize(sessionData, options);

                // 4. 将 JSON 字符串异步写入文件
                //System.IO.File.WriteAllText(filePath, jsonString);
                using (var writer = new StreamWriter(filePath, false))
                {
                    await writer.WriteAsync(jsonString); // 使用 WriteAsync 保持异步
                }

                Console.WriteLine($"JSON 文件已成功创建于: {filePath}");
                Console.WriteLine("文件内容:");
                Console.WriteLine(jsonString);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建 JSON 文件时发生错误: {ex.Message}");
            }
        }

        public static async Task<Settings> ReadSessionJsonAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"错误：文件未找到于 {filePath}");
                return null;
            }

            try
            {
                using (var reader = new StreamReader(filePath))
                {
                    // 1. 异步读取所有文本内容
                    string jsonString = await reader.ReadToEndAsync();

                    // 2. 将 JSON 字符串反序列化回 SessionInfo 对象
                    Settings sessionData = JsonSerializer.Deserialize<Settings>(jsonString);

                    if (sessionData != null)
                    {
                        Console.WriteLine("JSON 文件异步读取成功。");
                    }
                    return sessionData;
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"反序列化 JSON 时发生错误: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取文件时发生错误: {ex.Message}");
                return null;
            }
        }
    }
}
