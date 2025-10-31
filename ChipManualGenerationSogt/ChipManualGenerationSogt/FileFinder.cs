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

        // �����ı��ļ���չ�����ɸ�����Ҫ������
        public static readonly string[] DefaultTextExtensions = {
            ".txt",  ".csv", ".s2p", ".s3p", ".s4p"
        };

        /// <summary>
        /// ʹ��Ĭ���ı���չ����ʼ����������
        /// </summary>
        /// <param name="rootDirectory">Ҫ�����ĸ�Ŀ¼����Ϊ null ��գ���ʹ�õ�ǰ����Ŀ¼��</param>
        public TextFileFinder(string rootDirectory = null)
            : this(rootDirectory, DefaultTextExtensions)
        {
        }

        /// <summary>
        /// ʹ���Զ�����չ���б��ʼ����������
        /// </summary>
        /// <param name="rootDirectory">Ҫ�����ĸ�Ŀ¼��</param>
        /// <param name="extensions">Ҫ���ҵ��ļ���չ���б��� new[] { ".txt", ".log" }���������ִ�Сд��</param>
        public TextFileFinder(string rootDirectory, string[] extensions)
        {
            _rootDirectory = string.IsNullOrEmpty(rootDirectory)
                ? Directory.GetCurrentDirectory()
                : Path.GetFullPath(rootDirectory);

            if (!Directory.Exists(_rootDirectory))
                throw new DirectoryNotFoundException($"ָ���ĸ�Ŀ¼������: {_rootDirectory}");

            if (extensions == null || extensions.Length == 0)
                throw new ArgumentException("��չ���б���Ϊ�ա�", nameof(extensions));

            // ��׼����չ����ȷ���� '.' ��ͷ��ȫСд�����ڱȽ�
            _extensions = extensions
                .Select(ext => ext.StartsWith(".") ? ext.ToLowerInvariant() : "." + ext.ToLowerInvariant())
                .Distinct()
                .ToArray();
        }

        /// <summary>
        /// ��������ƥ����չ�����ļ���������������ڸ�Ŀ¼��·���б�
        /// </summary>
        /// <returns>����ĸ˳����������·��ֻ���б�</returns>
        public IReadOnlyList<string> FindAllTextFiles()
        {
            try
            {
                var allFiles = new List<string>();

                // ����ÿ����չ�����ռ��ļ�
                foreach (string ext in _extensions)
                {
                    allFiles.AddRange(Directory.GetFiles(_rootDirectory, "*" + ext, SearchOption.AllDirectories));
                }

                // ȥ�أ���ͬ��չ�������ص�����ͨ�����ᣩ
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
        /// ��ȡĿ��·������ڻ�׼·�������·����
        /// ���� .NET Framework �� .NET Core/5+��
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
        // �ʼ��
        public static AmpfilierFilesbyGroup ProcessFiles1(IEnumerable<string> allFilePaths)
        {
            var result = new AmpfilierFilesbyGroup();

            foreach (string item in allFilePaths)
            {
                string fullPath = @"CopiedReports\" + item;
                string fileName = Path.GetFileName(fullPath);

                // �ж��ļ����Ͳ�����
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
                string fullPath = @"CopiedReports\" + item;
                string fileName = Path.GetFileName(fullPath);

                // �ж��ļ����Ͳ�����
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


        // ---- �ļ������ж� ----
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

        // ---- ͨ�÷����߼� ----
        private static void AddToGroup(
            Dictionary<string, List<string>> byTemp,
            Dictionary<string, List<string>> byVD,
            string fullPath,
            string fileName)
        {
            // �Ƴ���չ��
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            string[] parts = baseName.Split('_');

            // ��ȡ�¶ȣ��� "deg" ��β�ĶΣ�
            string temperature = parts.FirstOrDefault(p => p.EndsWith("deg", StringComparison.OrdinalIgnoreCase))
                               ?? "UnknownTemp";

            // ��ȡ�������������� "VD=" �ĶΣ�
            //string temp = "";
            string temp = parts.FirstOrDefault(p => p.Contains("VD="))
                             ?? "UnknownParam";
            string elecParam= temp.Split('&')[0];
            // ���¶ȷ���
            if (!byTemp.TryGetValue(temperature, out var tempList))
            {
                tempList = new List<string>();
                byTemp[temperature] = tempList;
            }
            tempList.Add(fullPath);

            // ��������������
            if (!byVD.TryGetValue(elecParam, out var vdList))
            {
                vdList = new List<string>();
                byVD[elecParam] = vdList;
            }
            vdList.Add(fullPath);
        }
    }
}
