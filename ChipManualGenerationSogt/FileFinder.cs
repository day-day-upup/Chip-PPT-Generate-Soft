using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChipManualGenerationSogt
{

  
    public class TextFileFinder
    {
        private readonly string _rootDirectory;
        private readonly string[] _extensions;

        // 常见文本文件扩展名（可根据需要调整）
        public static readonly string[] DefaultTextExtensions = {
            ".txt",  ".csv", ".s2p", ".s3p", ".s4p"
        };

        /// <summary>
        /// 使用默认文本扩展名初始化查找器。
        /// </summary>
        /// <param name="rootDirectory">要搜索的根目录。若为 null 或空，则使用当前工作目录。</param>
        public TextFileFinder(string rootDirectory = null)
            : this(rootDirectory, DefaultTextExtensions)
        {
        }

        /// <summary>
        /// 使用自定义扩展名列表初始化查找器。
        /// </summary>
        /// <param name="rootDirectory">要搜索的根目录。</param>
        /// <param name="extensions">要查找的文件扩展名列表（如 new[] { ".txt", ".log" }），不区分大小写。</param>
        public TextFileFinder(string rootDirectory, string[] extensions)
        {
            _rootDirectory = string.IsNullOrEmpty(rootDirectory)
                ? Directory.GetCurrentDirectory()
                : Path.GetFullPath(rootDirectory);

            if (!Directory.Exists(_rootDirectory))
                throw new DirectoryNotFoundException($"指定的根目录不存在: {_rootDirectory}");

            if (extensions == null || extensions.Length == 0)
                throw new ArgumentException("扩展名列表不能为空。", nameof(extensions));

            // 标准化扩展名：确保以 '.' 开头且全小写，便于比较
            _extensions = extensions
                .Select(ext => ext.StartsWith(".") ? ext.ToLowerInvariant() : "." + ext.ToLowerInvariant())
                .Distinct()
                .ToArray();
        }

        /// <summary>
        /// 查找所有匹配扩展名的文件，并返回其相对于根目录的路径列表。
        /// </summary>
        /// <returns>按字母顺序排序的相对路径只读列表。</returns>
        public IReadOnlyList<string> FindAllTextFiles()
        {
            try
            {
                var allFiles = new List<string>();

                // 遍历每个扩展名并收集文件
                foreach (string ext in _extensions)
                {
                    allFiles.AddRange(Directory.GetFiles(_rootDirectory, "*" + ext, SearchOption.AllDirectories));
                }

                // 去重（不同扩展名可能重叠，但通常不会）
                var uniqueFullPaths = allFiles.Distinct().ToArray();

                var relativePaths = new List<string>(uniqueFullPaths.Length);
                foreach (string fullPath in uniqueFullPaths)
                {
                    string relativePath = GetRelativePath(_rootDirectory, fullPath);
                    relativePaths.Add(relativePath);
                }

                relativePaths.Sort(StringComparer.OrdinalIgnoreCase);
                return relativePaths.AsReadOnly();
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException)
            {
                throw new InvalidOperationException($"Access File Error: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 获取目标路径相对于基准路径的相对路径。
        /// 兼容 .NET Framework 和 .NET Core/5+。
        /// </summary>
        private static string GetRelativePath(string basePath, string targetPath)
        {
#if NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
            return Path.GetRelativePath(basePath, targetPath);
#else
            Uri baseUri = new Uri(basePath.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar);
            Uri targetUri = new Uri(targetPath);
            Uri relativeUri = baseUri.MakeRelativeUri(targetUri);
            return Uri.UnescapeDataString(relativeUri.ToString()).Replace('/', Path.DirectorySeparatorChar);
#endif
        }
    }

public static class AmplifierFileProcessor
    {
        // 最开始的
        public static AmpfilierFilesbyGroup ProcessFiles1(IEnumerable<string> allFilePaths)
        {
            var result = new AmpfilierFilesbyGroup();

            foreach (string item in allFilePaths)
            {
                string fullPath = @"CopiedReports\" + item;
                string fileName = Path.GetFileName(fullPath);

                // 判断文件类型并处理
                if (IsDataSparaFile(fileName))
                {
                    result.DataSparaFilePaths.Add(fullPath);
                    AddToGroup(result.DataSparabyTemp, result.DataSparabyVD, fullPath, fileName);
                }
                else if (IsNFFile(fileName))
                {
                    result.NFFiles.Add(fullPath);
                    AddToGroup(result.NFbyTemp, result.NFbyVD, fullPath, fileName);
                }
                else if (IsPsatFile(fileName))
                {
                    result.PsatFiles.Add(fullPath);
                    AddToGroup(result.PsatbyTemp, result.PsatbyVD, fullPath, fileName);
                }
                else if (IsOIP3File(fileName))
                {
                    result.OIP3Files.Add(fullPath);
                    AddToGroup(result.OIP3byTemp, result.OIP3byVD, fullPath, fileName);
                }
                else if (IsPxdBFile(fileName))
                {
                    result.PxdBFiles.Add(fullPath);
                    AddToGroup(result.PxdBbyTemp, result.PxdBbyVD, fullPath, fileName);
                }
            }

            return result;
        }


        public static AmpfilierFilesbyGroup ProcessFiles(IEnumerable<string> allFilePaths)
        {
            var result = new AmpfilierFilesbyGroup();

            foreach (string item in allFilePaths)
            {
                string fullPath = System.IO.Path.Combine(Global.TempBasePath, item);
                string fileName = Path.GetFileName(fullPath);

                // 判断文件类型并处理
                if (IsDataSparaFile(fileName))
                {
                    result.DataSparaFilePaths.Add(fullPath);
                    AddToGroup(result.DataSparabyTemp, result.DataSparabyVD, fullPath, fileName);
                }
                else if (IsNFFile(fileName))
                {
                    result.NFFiles.Add(fullPath);
                    AddToGroup(result.NFbyTemp, result.NFbyVD, fullPath, fileName);
                }
                else if (IsPsatFile(fileName))
                {
                    result.PsatFiles.Add(fullPath);
                    AddToGroup(result.PsatbyTemp, result.PsatbyVD, fullPath, fileName);
                }
                else if (IsOIP3File(fileName))
                {
                    result.OIP3Files.Add(fullPath);
                    AddToGroup(result.OIP3byTemp, result.OIP3byVD, fullPath, fileName);
                }
                else if (IsPxdBFile(fileName))
                {
                    result.PxdBFiles.Add(fullPath);
                    AddToGroup(result.PxdBbyTemp, result.PxdBbyVD, fullPath, fileName);
                }
            }

            return result;
        }


        // ---- 文件类型判断 ----
        private static bool IsDataSparaFile(string fileName) =>
            fileName.EndsWith("DataSpara.s2p", StringComparison.OrdinalIgnoreCase);

        private static bool IsNFFile(string fileName) =>
            fileName.EndsWith("NF.txt", StringComparison.OrdinalIgnoreCase);

        private static bool IsPsatFile(string fileName) =>
            fileName.EndsWith("Psat.txt", StringComparison.OrdinalIgnoreCase);

        private static bool IsOIP3File(string fileName) =>
            fileName.EndsWith("OIP3 vs. Frequency.txt", StringComparison.OrdinalIgnoreCase);

        private static bool IsPxdBFile(string fileName) =>
            fileName.EndsWith("PxdB vs. Frequency.txt", StringComparison.OrdinalIgnoreCase);

        // ---- 通用分组逻辑 ----
        private static void AddToGroup(
            Dictionary<string, List<string>> byTemp,
            Dictionary<string, List<string>> byVD,
            string fullPath,
            string fileName)
        {
            // 移除扩展名
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            string[] parts = baseName.Split('_');

            // 提取温度（以 "deg" 结尾的段）
            string temperature = parts.FirstOrDefault(p => p.EndsWith("deg", StringComparison.OrdinalIgnoreCase))
                               ?? "UnknownTemp";

            // 提取电气参数（包含 "VD=" 的段）
            //string temp = "";
            string temp = parts.FirstOrDefault(p => p.Contains("VD="))
                             ?? "UnknownParam";
            string elecParam= temp.Split('&')[0];
            // 按温度分组
            if (!byTemp.TryGetValue(temperature, out var tempList))
            {
                tempList = new List<string>();
                byTemp[temperature] = tempList;
            }
            tempList.Add(fullPath);

            // 按电气参数分组
            if (!byVD.TryGetValue(elecParam, out var vdList))
            {
                vdList = new List<string>();
                byVD[elecParam] = vdList;
            }
            vdList.Add(fullPath);
        }
    }
}
