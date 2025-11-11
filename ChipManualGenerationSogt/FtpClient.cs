using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ChipManualGenerationSogt
{
    public class FtpClient
    {
        static private readonly string _ftpServer = "ftp://192.168.1.209:12315";// 后续用配置文件配置
        //static private readonly NetworkCredential _credential = new NetworkCredential("yp", "123456");

        //private readonly string _ftpServer;
        static private readonly string _user="sa";
        static private readonly string _password= "qotana";

        //public FtpUploader(string ftpServer, string user, string password)
        //{
        //    _ftpServer = ftpServer.TrimEnd('/');
        //    _user = user;
        //    _password = password;
        //}

        /// <summary>
        /// 异步上传整个文件夹
        /// </summary>
        /// <summary>
        /// 异步上传整个文件夹，返回是否成功
        /// </summary>
        public static async Task<bool> UploadFolderAsync(string localFolder, string remoteFolder)
        {
            if (!Directory.Exists(localFolder))
                throw new DirectoryNotFoundException($"本地文件夹不存在: {localFolder}");

            await CreateRemoteDirectoryAsync(remoteFolder);

            List<string> failedFiles = new List<string>();

            // 上传文件
            foreach (var file in Directory.GetFiles(localFolder))
            {
                var fileName = Path.GetFileName(file);
                bool success = await UploadFileAsync(file, $"{remoteFolder}/{fileName}");
                if (!success)
                    failedFiles.Add(file);
            }

            // 递归上传子目录
            foreach (var subDir in Directory.GetDirectories(localFolder))
            {
                var dirName = Path.GetFileName(subDir);
                bool success = await UploadFolderAsync(subDir, $"{remoteFolder}/{dirName}");
                if (!success)
                    failedFiles.Add(subDir);
            }

            if (failedFiles.Count > 0)
            {
                Console.WriteLine("? 以下文件/目录上传失败：");
                foreach (var f in failedFiles)
                    Console.WriteLine($"   {f}");
                return false;
            }

            Console.WriteLine($"? 文件夹上传成功: {localFolder}");
            return true;
        }

        /// <summary>
        /// 异步上传单个文件，返回是否成功
        /// </summary>
        public static async Task<bool> UploadFileAsync(string localFile, string remoteFile)
        {
            string uri = $"{_ftpServer}/{remoteFile}".Replace("\\", "/");
            Console.WriteLine($"?? 开始上传: {localFile}");

            try
            {
                var request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential(_user, _password);
                request.UseBinary = true;
                request.UsePassive = true;

                using (var fileStream = File.OpenRead(localFile))
                using (var ftpStream = await request.GetRequestStreamAsync())
                {
                    byte[] buffer = new byte[81920];
                    int bytesRead;
                    long totalBytes = fileStream.Length;
                    long uploaded = 0;

                    while ((bytesRead = await fileStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                    {
                        await ftpStream.WriteAsync(buffer, 0, bytesRead);
                        uploaded += bytesRead;

                        Console.Write($"\r    进度: {uploaded * 100.0 / totalBytes:F1}%");
                    }
                }

                using (var response = (FtpWebResponse)await request.GetResponseAsync())
                {
                    Console.WriteLine($"\r? 上传完成: {Path.GetFileName(localFile)} ({response.StatusDescription.Trim()})");
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n? 上传失败: {localFile}\n   原因: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 异步创建远程目录（如果不存在）
        /// </summary>
        private static async Task CreateRemoteDirectoryAsync(string remoteDir)
        {
            string[] parts = remoteDir.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            string currentPath = "";

            foreach (var part in parts)
            {
                currentPath += "/" + part;
                string uri = $"{_ftpServer}{currentPath}";

                try
                {
                    var request = (FtpWebRequest)WebRequest.Create(uri);
                    request.Method = WebRequestMethods.Ftp.MakeDirectory;
                    request.Credentials = new NetworkCredential(_user, _password);
                    request.UsePassive = true;
                    request.UseBinary = true;

                    var response = (FtpWebResponse)await request.GetResponseAsync();
                }
                catch (WebException ex)
                {
                    // 若目录已存在则忽略 550 错误
                    if (!ex.Message.Contains("550"))
                        throw;
                }
            }
        }


        /// <summary>
        /// 异步下载整个文件夹
        /// </summary>
        static public async Task<bool> DownloadFolderAsync(string remoteFolder, string localFolder)
        {
            try
            {
                Directory.CreateDirectory(localFolder);

                var entries = await ListDirectoryDetailsAsync(remoteFolder);
                foreach (var entry in entries)
                {
                    string name = entry.Name;
                    string remotePath = $"{remoteFolder}/{name}";
                    string localPath = Path.Combine(localFolder, name);

                    if (entry.IsDirectory)
                    {
                        await DownloadFolderAsync(remotePath, localPath);
                    }
                    else
                    {
                        await DownloadFileAsync(remotePath, localPath);
                    }
                }

                Console.WriteLine($"? 下载完成: {remoteFolder}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"? 下载失败: {remoteFolder}\n原因: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 异步下载单个文件
        /// </summary>
        static public async Task DownloadFileAsync(string remoteFile, string localFile)
        {
            string uri = BuildFtpUri(_ftpServer, remoteFile);
            //string uri = $"{_ftpServer}/{remoteFile}".Replace("\\", "/");
            Console.WriteLine($"?? 正在下载文件: {uri}");

            var request = (FtpWebRequest)WebRequest.Create(uri);
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            request.Credentials = new NetworkCredential(_user, _password);
            request.UseBinary = true;
            request.UsePassive = true;

            using (var response = (FtpWebResponse)await request.GetResponseAsync())
            using (var responseStream = response.GetResponseStream())
            using (var fileStream = File.Create(localFile))

            await responseStream.CopyToAsync(fileStream);

            Console.WriteLine($"   ? 下载完成: {Path.GetFileName(localFile)}");
        }

        /// <summary>
        /// 异步下载单个文件，返回是否成功ex 表示的是扩展
        static public async Task<bool> DownloadFileAsyncEx(string remoteFile, string localFile)
        {
            try
            {
                string uri = BuildFtpUri(_ftpServer, remoteFile);
                Console.WriteLine($"?? 正在下载文件: {uri}");

                var request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.DownloadFile;
                request.Credentials = new NetworkCredential(_user, _password);
                request.UseBinary = true;
                request.UsePassive = true;

                using (var response = (FtpWebResponse)await request.GetResponseAsync())
                using (var responseStream = response.GetResponseStream())
                using (var fileStream = File.Create(localFile))
                {
                    await responseStream.CopyToAsync(fileStream);
                }

                Console.WriteLine($"   ? 下载完成: {Path.GetFileName(localFile)}");
                return true; // 下载成功
            }
            catch (WebException webEx)
            {
                // 捕获与网络或 FTP 相关的错误
                Console.WriteLine($"!!! 文件下载失败 ({remoteFile}): {webEx.Message}");

                // 尝试获取更具体的 FTP 状态码
                if (webEx.Response is FtpWebResponse ftpResponse)
                {
                    Console.WriteLine($"!!! FTP 状态: {ftpResponse.StatusCode} - {ftpResponse.StatusDescription}");
                }

                // 如果文件创建失败（例如，路径无效），也会抛出异常，此时 localFile 可能不存在
                return false;
            }
            catch (Exception ex)
            {
                // 捕获其他非网络错误（如文件系统权限问题等）
                Console.WriteLine($"!!! 文件下载发生意外错误 ({remoteFile}): {ex.Message}");
                return false;
            }
        }
        static private async Task<List<FtpEntry>> ListDirectoryDetailsAsync(string remoteFolder)
        {
            string uri = BuildFtpUri(_ftpServer, remoteFolder);

            var request = (FtpWebRequest)WebRequest.Create(uri);
            request.Method = WebRequestMethods.Ftp.ListDirectoryDetails; // ← LIST 命令
            request.Credentials = new NetworkCredential(_user, _password);
            request.UsePassive = true;  // Windows FTP 通常用 Passive=true
            request.UseBinary = true;

            var result = new List<FtpEntry>();

            using (var response = (FtpWebResponse)await request.GetResponseAsync())
            using (var stream = response.GetResponseStream())
            using (var reader = new StreamReader(stream))
            {
                while (!reader.EndOfStream)
                {
                    string line = await reader.ReadLineAsync();
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var entry = ParseFtpEntry(line);
                    if (entry != null && !string.IsNullOrEmpty(entry.Name))
                    {
                        result.Add(entry);
                    }
                }
            }

            return result;
        }
        class FtpEntry
        {
            public string Name { get; set; }
            public bool IsDirectory { get; set; }
        }


        /// <summary>
        /// 解析 FTP LIST 返回的每一行
        /// 兼容 Unix 风格（大多数 FTP 服务器）
        /// </summary>
        static private FtpEntry ParseFtpEntry(string line)
        {
            // 尝试 Unix 风格（以 d 或 - 开头）
            if (line.Length >= 10 && (line[0] == 'd' || line[0] == '-'))
            {
                var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 9)
                {
                    string name = string.Join(" ", parts.Skip(8));
                    return new FtpEntry { Name = name, IsDirectory = line[0] == 'd' };
                }
            }

            // Windows 风格：日期 时间 [DIR或大小] 名称
            // 例如：
            // 01-09-25  03:37PM       <DIR>          FolderName
            // 01-09-25  03:37PM                12345 FileName with spaces.txt
            if (line.Length > 30 && char.IsDigit(line[0]))
            {
                // 找到最后一个 "      "（多个空格）之后的部分
                // 实际名称从第 38 列左右开始（但不固定），更可靠方式：找 <DIR> 或数字后的内容
                int dirIndex = line.IndexOf("<DIR>", StringComparison.OrdinalIgnoreCase);
                if (dirIndex >= 0)
                {
                    // 是目录
                    string name = line.Substring(dirIndex + 5).Trim();
                    return new FtpEntry { Name = name, IsDirectory = true };
                }
                else
                {
                    // 尝试找文件大小（一串数字）
                    // 从日期时间之后跳过，找第一个数字字段（文件大小），然后后面是文件名
                    // 更简单：从行末往前找第一个非空格段？但不可靠
                    // 更健壮：按固定格式切分（Windows LIST 通常前38字符是日期+时间+大小/DIR）
                    if (line.Length > 38)
                    {
                        string name = line.Substring(38).Trim();
                        if (!string.IsNullOrEmpty(name))
                        {
                            return new FtpEntry { Name = name, IsDirectory = false };
                        }
                    }
                }
            }

            // 如果无法解析，保守认为是文件（避免误判为目录）
            // 或记录警告
            Console.WriteLine($"?? 无法解析 FTP 行: {line}");
            return new FtpEntry { Name = line.Trim(), IsDirectory = false };
        }
        private static async Task<bool> IsDirectoryAsync(string path)
        {
            try
            {
                var request = (FtpWebRequest)WebRequest.Create(path);
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(_user, _password);
                request.UsePassive = false;
                request.UseBinary = true;

                using (var response = (FtpWebResponse)await request.GetResponseAsync())
                {
                    // 能列出内容就是目录
                    return true;
                }
            }
            catch (WebException ex)
            {
                // 550 错误一般代表不是目录，是文件
                if (ex.Response is FtpWebResponse ftpResponse &&
                    ftpResponse.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                    return false;

                throw;
            }
        }

        private static string BuildFtpUri(string baseUri, string relativePath)
        {
            // 针对每一级目录分别编码，防止 &、空格、中文等破坏路径
            var parts = relativePath.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var encodedParts = parts.Select(Uri.EscapeDataString);
            return $"{baseUri.TrimEnd('/')}/{string.Join("/", encodedParts)}";
        }
    }
}
