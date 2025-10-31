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
            // 1. ʵ���������������
           

            // 2. �������л�ѡ�� (��ѡ, ֻ��Ϊ���������)
            var options = new JsonSerializerOptions
            {
                // ʹ JSON �����ʽ�� (�������ͻ���)
                WriteIndented = true
            };

            try
            {
                // 3. ���л�����Ϊ JSON �ַ���
                string jsonString = JsonSerializer.Serialize(sessionData, options);

                // 4. �� JSON �ַ����첽д���ļ�
                //System.IO.File.WriteAllText(filePath, jsonString);
                using (var writer = new StreamWriter(filePath, false))
                {
                    await writer.WriteAsync(jsonString); // ʹ�� WriteAsync �����첽
                }

                Console.WriteLine($"JSON �ļ��ѳɹ�������: {filePath}");
                Console.WriteLine("�ļ�����:");
                Console.WriteLine(jsonString);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"���� JSON �ļ�ʱ��������: {ex.Message}");
            }
        }

        public static async Task<Settings> ReadSessionJsonAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"�����ļ�δ�ҵ��� {filePath}");
                return null;
            }

            try
            {
                using (var reader = new StreamReader(filePath))
                {
                    // 1. �첽��ȡ�����ı�����
                    string jsonString = await reader.ReadToEndAsync();

                    // 2. �� JSON �ַ��������л��� SessionInfo ����
                    Settings sessionData = JsonSerializer.Deserialize<Settings>(jsonString);

                    if (sessionData != null)
                    {
                        Console.WriteLine("JSON �ļ��첽��ȡ�ɹ���");
                    }
                    return sessionData;
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"�����л� JSON ʱ��������: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"��ȡ�ļ�ʱ��������: {ex.Message}");
                return null;
            }
        }
    }
}
