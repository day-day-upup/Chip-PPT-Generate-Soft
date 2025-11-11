using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.DataSourceVersioning;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.IO;
using System.Text.Json.Serialization;
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

    #region 参数表格的数据模型
    public class ParameterItem
    {
        // 使用 JsonPropertyName 属性来处理 Key/Value 属性名
        [JsonPropertyName("Key")]
        public string Key { get; set; }

        [JsonPropertyName("Value")]
        public string Value { get; set; }
    }

    public class SupplyCurrentItem
    {
        [JsonPropertyName("VD")]
        public string VD { get; set; }

        [JsonPropertyName("VG")]
        public string VG { get; set; }

        [JsonPropertyName("Current (mA)")]
        public string Current_mA { get; set; } // C# 属性名不能有空格，使用 JsonPropertyName 映射
    }

    public class DescriptionItem1
    {
        [JsonPropertyName("Component")]
        public string Component { get; set; }

        [JsonPropertyName("Description")]
        public string Description { get; set; }
    }

    public class DescriptionItem2
    {
        [JsonPropertyName("Pin")]
        public string Pin { get; set; }

        [JsonPropertyName("Function")]
        public string Function { get; set; }

        [JsonPropertyName("Detail")]
        public string Detail { get; set; }
    }
    public class PerformanceGroup
    {
        // 注意：您 ViewModel 中的 Type 属性，在 JSON 术语中通常对应 Typ (Typical)
        [JsonPropertyName("Min")]
        public string Min { get; set; } = string.Empty;

        [JsonPropertyName("Typ")] // 映射您 ViewModel 中的 Type 属性
        public string Typ { get; set; } = string.Empty;

        [JsonPropertyName("Max")]
        public string Max { get; set; } = string.Empty;
    }

    public class ParameterDetail
    {
        [JsonPropertyName("Name")]
        public string Name { get; set; }

        [JsonPropertyName("Unit")]
        public string Unit { get; set; }

        [JsonPropertyName("Groups")]
        public List<PerformanceGroup> Groups { get; set; }
    }
    public class ParamTables
    {
        [JsonPropertyName("BasicParameters")]
        public List<ParameterItem> BasicParameters { get; set; }

        [JsonPropertyName("FeatureParameters")]
        public List<ParameterItem> FeatureParameters { get; set; }
        [JsonPropertyName("DetailedPerformance")]
        public List<ParameterDetail> DetailedPerformance { get; set; }

        [JsonPropertyName("AbsoluteMaximumRatings")]
        public List<ParameterItem> AbsoluteMaximumRatings { get; set; }

        [JsonPropertyName("SupplyCurrentVdVg")]
        public List<SupplyCurrentItem> SupplyCurrentVdVg { get; set; }

        [JsonPropertyName("Notes")]
        public List<string> Notes { get; set; } // Notes 是一个字符串数组

        [JsonPropertyName("Description1")]
        public List<DescriptionItem1> Description1 { get; set; }

        [JsonPropertyName("Description2")]
        public List<DescriptionItem2> Description2 { get; set; }

        [JsonPropertyName("TurnOnProcedure")]
        public List<string> TurnOnProcedure { get; set; }

        [JsonPropertyName("TurnOffProcedure")]
        public List<string> TurnOffProcedure { get; set; }
    }

    // 顶级（Root）对象
    public class ProductData
    {
        [JsonPropertyName("ProductName")]
        public string ProductName { get; set; }

        [JsonPropertyName("ModelNumber")]
        public string ModelNumber { get; set; }

        [JsonPropertyName("Tables")]
        public ParamTables Tables { get; set; }
    }

    #endregion


    public class DatabaseConnection
    {
        // [JsonPropertyName] 属性可选，但推荐用于确保 JSON 属性名和 C# 属性名一致

        [JsonPropertyName("Server")]
        public string Server { get; set; } = string.Empty;

        [JsonPropertyName("Database")]
        public string Database { get; set; } = string.Empty;

        [JsonPropertyName("UserID")]
        public string UserId { get; set; } = string.Empty; // 使用 UserId 命名更符合 C# 规范

        [JsonPropertyName("Password")]
        public string Password { get; set; } = string.Empty;

        [JsonPropertyName("Encrypt")]
        public bool Encrypt { get; set; }

        [JsonPropertyName("TrustServerCertificate")]
        public bool TrustServerCertificate { get; set; }

        /// <summary>
        /// 辅助方法：将此对象转换为 SQL Server 连接字符串
        /// </summary>
        public string ToConnectionString()
        {
            // 使用 StringBuilder 或 string.Format 构造连接字符串，确保格式正确
            return $"Server={Server};Database={Database};User ID={UserId};Password={Password};Encrypt={Encrypt};TrustServerCertificate={TrustServerCertificate};";
        }
    }


    public class DatabaseManager
    {
        // JSON 文件路径（示例）
        static public  string ConfigFilePath = System.IO.Path.Combine(Global.FileBasePath, "cfg.json");

        /// <summary>
        /// 将 DatabaseConnection 对象序列化为 JSON 字符串并异步保存到文件。
        /// </summary>
        /// <param name="connection">要保存的配置对象。</param>
        /// <returns>表示异步操作的 Task。</returns>
        public static async Task CreateConfigJsonFile(DatabaseConnection connection)
        {
            // 1. 配置序列化选项 (使 JSON 输出格式化)
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            try
            {
                // 2. 序列化对象为 JSON 字符串
                string jsonString = JsonSerializer.Serialize(connection, options);

                // 3. 将 JSON 字符串异步写入文件
                // 使用 StreamWriter 和 WriteAsync 是异步写入文件的一种可靠方式
                using (var writer = new StreamWriter(ConfigFilePath, false))
                {
                    await writer.WriteAsync(jsonString);
                }

                Console.WriteLine($"JSON 配置已成功创建于: {ConfigFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建 JSON 配置文件时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 从 JSON 文件中异步读取并反序列化为 DatabaseConnection 对象。
        /// </summary>
        /// <returns>反序列化后的 DatabaseConnection 对象，失败则返回 null。</returns>
        public static async Task<DatabaseConnection> ReadConfigJsonAsync()
        {
            if (!File.Exists(ConfigFilePath))
            {
                Console.WriteLine($"错误：配置文件未找到于 {ConfigFilePath}");
                // 如果文件不存在，返回一个空的或默认的连接对象，或者 null
                return null;
            }

            try
            {
                using (var reader = new StreamReader(ConfigFilePath))
                {
                    // 1. 异步读取所有文本内容
                    string jsonString = await reader.ReadToEndAsync();

                    // 2. 将 JSON 字符串反序列化回 DatabaseConnection 对象
                    // 注意：在 C# 8.0+ 的项目中，您应该使用 Task<DatabaseConnection?>
                    DatabaseConnection connection = JsonSerializer.Deserialize<DatabaseConnection>(jsonString);

                    if (connection != null)
                    {
                        Console.WriteLine("JSON 配置文件异步读取成功。");
                    }
                    return connection;
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
