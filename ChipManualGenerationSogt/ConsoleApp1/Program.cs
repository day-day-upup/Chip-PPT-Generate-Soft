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

            // ��ѡ���Զ�����־���������д���ļ���
            // copier.Log = msg => File.AppendAllText("copy.log", $"{DateTime.Now:HH:mm:ss} {msg}\n");

            //try
            //{
            //    copier.CopyMatchingSubFolders(
            //        networkRoot: @"\\DATAPC03\RFAutoTestReport$\Chip Verification",
            //        searchPattern: "L004x",
            //        localTargetBase: Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiedReports"),
            //        username: "",   // ���ձ�ʾʹ�õ�ǰ�û�ƾ��
            //        password: ""
            //    );
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"?? �����쳣: {ex.Message}");
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
        // �ݹ鸴������Ŀ¼
        static void CopyDirectory(string sourceDir, string destDir, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDir);
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"ԴĿ¼�����ڻ��޷�����: {sourceDir}");

            // ���Ŀ�겻���ڣ�����
            Directory.CreateDirectory(destDir);

            // �����ļ�
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDir, file.Name);
                file.CopyTo(temppath, overwrite: false); // �����������ļ�
            }

            // ������Ŀ¼
            if (copySubDirs)
            {
                DirectoryInfo[] dirs = dir.GetDirectories();
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDir, subdir.Name);
                    // �ݹ�
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
                Console.WriteLine($"?? �޷����ӹ���·����������: {result}");
            }
        }

        // �����ļ���ƥ�����1
        static void test1()
        {
            string networkPath = @"\\DATAPC03\RFAutoTestReport$\Chip Verification";
            string username = ""; // �粻��Ҫ��֤��������
            string password = "";
            string searchPattern = "L004x";

            // ���ӹ����ļ��У��������Ҫ��֤Ҳ���Ե��ã����û���/�������գ�
             ConnectToShare(networkPath, username, password);
            //if (connResult != 0)
            //{
            //    Console.WriteLine($"?? �޷����ӵ�����·���������� {connResult}����������ֱ�ӷ��ʡ�");
            //}

            try
            {
                // ��������·���µ�һ��Ŀ¼����Ŀ¼��
                var parentDirs = Directory.GetDirectories(networkPath, "*", SearchOption.TopDirectoryOnly);
                foreach (var parent in parentDirs)
                {
                    Console.WriteLine($"��鸸Ŀ¼: {parent}");

                    // ��ÿ����Ŀ¼�²������ļ��������� searchPattern ����
                    var childDirs = Directory.GetDirectories(parent, "*", SearchOption.TopDirectoryOnly);
                    foreach (var child in childDirs)
                    {
                        if (child.IndexOf(searchPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"?? �ҵ�ƥ�����ļ���: {child}");

                            // Ŀ��·������ǰ����Ŀ¼�µ�ͬ���ļ��У��Զ������ͻ��
                            string destRoot = Environment.CurrentDirectory;
                            string destFolderName = Path.GetFileName(child);
                            string destPath = Path.Combine(destRoot, destFolderName);

                            destPath = MakeUniqueDirectory(destPath);

                            Console.WriteLine($"?? ���Ƶ�: {destPath}");
                            CopyDirectory(child, destPath, true);
                            Console.WriteLine("   ? �������");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("? ���ʻ��Ƴ���: " + ex.ToString());
            }
            finally
            {
                // �Ͽ��������ӣ������׳��쳣��
                DisconnectFromShare(networkPath);
            }
        }

        // ���Ŀ��·���Ѵ��ڣ�����һ������ͻ����·������� _copy_n �� ʱ�����
        static string MakeUniqueDirectory(string path)
        {
            if (!Directory.Exists(path))
                return path;

            // ���������ֺ�׺
            int i = 1;
            string basePath = path;
            while (Directory.Exists(path))
            {
                path = $"{basePath}_copy{i}";
                i++;
                // Ϊ��������ѭ������ i �ܴ�Ҳ���Ը�Ϊʱ�������
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
                    // ��ȫ��ȡ�ɿ��ֶΣ�PersonID��StoreID��TerritoryID ����Ϊ NULL��
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
                Console.WriteLine($"��ȡ {analyzer.S11.Count} �� S11 ���ݵ�");
                foreach (var s in analyzer.S11.Take(20))
                {
                    Console.WriteLine($"Freq: {s.FreqGHz:F3} GHz, S11: {s.DbValue:F2} dB");
                }
            }
            else
            {
                Console.WriteLine("��ȡʧ�ܣ�");
            }
        }



    }

    public class S_Parameters
    {
        public double FreqGHz { get; set; }   // Ƶ�ʣ�GHz��
        public double DbValue { get; set; }   // ���ȣ�dB��
        public double PhaseValue { get; set; } // ��λ���ȣ�
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
            // ����У��
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("�ļ�·������Ϊ��", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("S2P �ļ�δ�ҵ�", filePath);
            try
            {
                // �����������
                _s11.Clear();
                _s21.Clear();
                _s12.Clear();
                _s22.Clear();
                foreach (var line in File.ReadLines(filePath))
                {
                    var trimmedLine = line.Trim();

                    // �������к�ע���У��� ! �� # ��ͷ��
                    if (string.IsNullOrEmpty(trimmedLine) ||
                        trimmedLine.StartsWith('!') ||
                        trimmedLine.StartsWith('#'))
                    {
                        continue;
                    }

                    // ���հ��ַ��ָ֧�ֿո��Ʊ���ȣ�
                    var tokens = trimmedLine.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);

                    // S2P ��׼��Ӧ�� 9 ����ֵ��Ƶ�� + 8 �� S ������
                    if (tokens.Length < 9)
                        continue; // �� throw �쳣����������

                    if (!double.TryParse(tokens[0], out double freqHz))
                        continue;

                    double freqGHz = freqHz / 1e9;

                    // ��ȫ������ֵ�������쳣�жϣ�
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
                        // ��ѡ����¼��־��������Ч��
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
                // �ɼ�¼��־��Logger?.LogError(ex, "��ȡ S2P �ļ�ʧ��");
                return false;
            }
        }


    }




}

