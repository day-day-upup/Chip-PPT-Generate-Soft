using System;
using QC=Microsoft.Data.SqlClient;
using DT = System.Data;
using System.Data;
using ProofOfConcept_SQL_CSharp;
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Runtime.InteropServices;
using ConsoleApp1;

namespace ProofOfConcept_SQL_CSharp
{
    public class Program
    {
        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(ref NETRESOURCE netResource,
       string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);

        [StructLayout(LayoutKind.Sequential)]
        private struct NETRESOURCE
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }
        static public void Main()
        {


            using (var connection = new QC.SqlConnection(
                "Server=192.168.1.77,1433;" +
                "Database=AdventureWorks2022;User ID=sa;" +
                "Password=123456;Encrypt=false;" +
                "TrustServerCertificate=True;Connection Timeout=30;"
                ))
            {
                try
                {
                    connection.Open();

                    Console.WriteLine("Connected successfully.");

                    Program.SelectRows(connection);

                    Console.WriteLine("Press any key to finish...");
                    Console.ReadKey(true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error connecting to database: " + ex.Message);
                }
            }


            //testS2ppaser(@"F:\PROJECT\ChipManualGeneration\KP414-11TAPE_RP2DR1-25_C10R6-8-VD=4.6V&ID=79mA_2025-06-12 19.33.41_25.0deg_SPara.s2p");

            var copier = new NetworkFolderCopier();

            // 可选：自定义日志输出（例如写入文件）
            // copier.Log = msg => File.AppendAllText("copy.log", $"{DateTime.Now:HH:mm:ss} {msg}\n");

            //try
            //{
            //    copier.CopyMatchingSubFolders(
            //        networkRoot: @"\\DATAPC03\RFAutoTestReport$\Chip Verification",
            //        searchPattern: "L004x",
            //        localTargetBase: Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiedReports"),
            //        username: "",   // 留空表示使用当前用户凭据
            //        password: ""
            //    );
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"?? 程序异常: {ex.Message}");
            //}

            string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";

            var repo = new TestRecordRepository(connStr);
            //var records = repo.GetRecordsByPN(
            //    pnList: new[] { "L004X"},
            //    startTime: new DateTime(2025, 10, 1),
            //    endTime: DateTime.Now
            //);
            var records = repo.GetRecordsByPN(
                keywords: new[] { "L004X" }
            );

            Console.WriteLine($"Found {records.Count} records.");
            foreach (var r in records)
            {
                Console.WriteLine($"{r.ID} | {r.PN} | {r.TestTime}");
            }
        }
        // 递归复制整个目录
        static void CopyDirectory(string sourceDir, string destDir, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDir);
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"源目录不存在或无法访问: {sourceDir}");

            // 如果目标不存在，创建
            Directory.CreateDirectory(destDir);

            // 复制文件
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDir, file.Name);
                file.CopyTo(temppath, overwrite: false); // 不覆盖已有文件
            }

            // 复制子目录
            if (copySubDirs)
            {
                DirectoryInfo[] dirs = dir.GetDirectories();
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDir, subdir.Name);
                    // 递归
                    CopyDirectory(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
        static void ConnectToShare(string networkPath, string username, string password)
        {
            NETRESOURCE nr = new NETRESOURCE
            {
                dwType = 1, // RESOURCETYPE_DISK
                lpRemoteName = networkPath
            };
            int result = WNetAddConnection2(ref nr, password, username, 0);

            if (result != 0)
            {
                Console.WriteLine($"?? 无法连接共享路径，错误码: {result}");
            }
        }

        // 共享文件夹匹配操作1
        static void test1()
        {
            string networkPath = @"\\DATAPC03\RFAutoTestReport$\Chip Verification";
            string username = ""; // 如不需要认证，可留空
            string password = "";
            string searchPattern = "L004x";

            // 连接共享文件夹（如果不需要认证也可以调用，但用户名/密码留空）
             ConnectToShare(networkPath, username, password);
            //if (connResult != 0)
            //{
            //    Console.WriteLine($"?? 无法连接到共享路径（错误码 {connResult}），将尝试直接访问。");
            //}

            try
            {
                // 遍历网络路径下的一级目录（父目录）
                var parentDirs = Directory.GetDirectories(networkPath, "*", SearchOption.TopDirectoryOnly);
                foreach (var parent in parentDirs)
                {
                    Console.WriteLine($"检查父目录: {parent}");

                    // 在每个父目录下查找子文件夹名包含 searchPattern 的项
                    var childDirs = Directory.GetDirectories(parent, "*", SearchOption.TopDirectoryOnly);
                    foreach (var child in childDirs)
                    {
                        if (child.IndexOf(searchPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"?? 找到匹配子文件夹: {child}");

                            // 目标路径：当前运行目录下的同名文件夹（自动避免冲突）
                            string destRoot = Environment.CurrentDirectory;
                            string destFolderName = Path.GetFileName(child);
                            string destPath = Path.Combine(destRoot, destFolderName);

                            destPath = MakeUniqueDirectory(destPath);

                            Console.WriteLine($"?? 复制到: {destPath}");
                            CopyDirectory(child, destPath, true);
                            Console.WriteLine("   ? 复制完成");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("? 访问或复制出错: " + ex.ToString());
            }
            finally
            {
                // 断开网络连接（无需抛出异常）
                DisconnectFromShare(networkPath);
            }
        }

        // 如果目标路径已存在，生成一个不冲突的新路径（添加 _copy_n 或 时间戳）
        static string MakeUniqueDirectory(string path)
        {
            if (!Directory.Exists(path))
                return path;

            // 尝试用数字后缀
            int i = 1;
            string basePath = path;
            while (Directory.Exists(path))
            {
                path = $"{basePath}_copy{i}";
                i++;
                // 为避免无限循环，若 i 很大也可以改为时间戳策略
                if (i > 1000)
                {
                    path = $"{basePath}_{DateTime.Now:yyyyMMddHHmmss}";
                    break;
                }
            }
            return path;
        }
        static void DisconnectFromShare(string networkPath)
        {
            WNetCancelConnection2(networkPath, 0, true);
        }
        static public void SelectRows(QC.SqlConnection connection)
        {
            using (var command = new QC.SqlCommand())
            {
                command.Connection = connection;
                command.CommandType = DT.CommandType.Text;
                command.CommandText = @"  
                SELECT TOP 5 *
                 FROM Sales.Customer; 
		         ";

                QC.SqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    // 安全读取可空字段（PersonID、StoreID、TerritoryID 可能为 NULL）
                    int? personId = reader.IsDBNull("PersonID") ? null : (int?)reader.GetInt32("PersonID");
                    int? storeId = reader.IsDBNull("StoreID") ? null : (int?)reader.GetInt32("StoreID");
                    int? territoryId = reader.IsDBNull("TerritoryID") ? null : (int?)reader.GetInt32("TerritoryID");

                    Console.WriteLine(
                        "{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}",
                        reader.GetInt32("CustomerID"),
                        personId?.ToString() ?? "NULL",
                        storeId?.ToString() ?? "NULL",
                        territoryId?.ToString() ?? "NULL",
                        reader.GetString("AccountNumber"),
                        reader.GetGuid("rowguid"),
                        reader.GetDateTime("ModifiedDate")
                    );
                }
            }
        }

        static void testS2ppaser(string filePath)
        {
            var analyzer = new S2PParser();
            bool success = analyzer.Parse(filePath);

            if (success)
            {
                Console.WriteLine($"读取 {analyzer.S11.Count} 个 S11 数据点");
                foreach (var s in analyzer.S11.Take(20))
                {
                    Console.WriteLine($"Freq: {s.FreqGHz:F3} GHz, S11: {s.DbValue:F2} dB");
                }
            }
            else
            {
                Console.WriteLine("读取失败！");
            }
        }



    }

    public class S_Parameters
    {
        public double FreqGHz { get; set; }   // 频率（GHz）
        public double DbValue { get; set; }   // 幅度（dB）
        public double PhaseValue { get; set; } // 相位（度）
    }
    public class S2PParser
    {
        private List<S_Parameters> _s11;

        public List<S_Parameters> S11 { get { return _s11; } }

        private List<S_Parameters> _s21;

        public List<S_Parameters> S21 { get { return _s21; } }

        private List<S_Parameters> _s12 ;

        public List<S_Parameters> S12 { get { return _s12; } }


        private List<S_Parameters> _s22 ;

        public List<S_Parameters> S22 { get { return _s22; } }


        public S2PParser()
        {
            _s11 = new List<S_Parameters>();
            _s21 = new List<S_Parameters>();
            _s12 = new List<S_Parameters>();
            _s22 = new List<S_Parameters>();
        }

        public S2PParser(string s2pFilePath)
        {
            _s11 = new List<S_Parameters>();
            _s21 = new List<S_Parameters>();
            _s12 = new List<S_Parameters>();
            _s22 = new List<S_Parameters>();
            Parse(s2pFilePath);
        }
        public bool Parse(string filePath)
        {
            // 参数校验
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("文件路径不能为空", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("S2P 文件未找到", filePath);
            try
            {
                // 清空现有数据
                _s11.Clear();
                _s21.Clear();
                _s12.Clear();
                _s22.Clear();
                foreach (var line in File.ReadLines(filePath))
                {
                    var trimmedLine = line.Trim();

                    // 跳过空行和注释行（以 ! 或 # 开头）
                    if (string.IsNullOrEmpty(trimmedLine) ||
                        trimmedLine.StartsWith('!') ||
                        trimmedLine.StartsWith('#'))
                    {
                        continue;
                    }

                    // 按空白字符分割（支持空格、制表符等）
                    var tokens = trimmedLine.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);

                    // S2P 标准行应有 9 个数值（频率 + 8 个 S 参数）
                    if (tokens.Length < 9)
                        continue; // 或 throw 异常，根据需求

                    if (!double.TryParse(tokens[0], out double freqHz))
                        continue;

                    double freqGHz = freqHz / 1e9;

                    // 安全解析数值（避免异常中断）
                    bool TryParseToken(int index, out double value)
                    {
                        value = 0;
                        return double.TryParse(tokens[index], out value);
                    }

                    if (!TryParseToken(1, out double s11Db) ||
                        !TryParseToken(2, out double s11Phase) ||
                        !TryParseToken(3, out double s21Db) ||
                        !TryParseToken(4, out double s21Phase) ||
                        !TryParseToken(5, out double s12Db) ||
                        !TryParseToken(6, out double s12Phase) ||
                        !TryParseToken(7, out double s22Db) ||
                        !TryParseToken(8, out double s22Phase))
                    {
                        // 可选：记录日志或跳过无效行
                        continue;
                    }

                    S11.Add(new S_Parameters { FreqGHz = freqGHz, DbValue = s11Db, PhaseValue = s11Phase });
                    S21.Add(new S_Parameters { FreqGHz = freqGHz, DbValue = s21Db, PhaseValue = s21Phase });
                    S12.Add(new S_Parameters { FreqGHz = freqGHz, DbValue = s12Db, PhaseValue = s12Phase });
                    S22.Add(new S_Parameters { FreqGHz = freqGHz, DbValue = s22Db, PhaseValue = s22Phase });
                }

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // 可记录日志：Logger?.LogError(ex, "读取 S2P 文件失败");
                return false;
            }
        }


    }




}

