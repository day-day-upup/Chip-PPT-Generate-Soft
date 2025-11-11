using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace ConsoleApp1
{
   





    public class NetworkFolderCopier
    {

        public string TargetPath { get; set; }
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

        // 日志输出委托（可选）
        public Action<string> Log { get; set; } = msg => Console.WriteLine(msg);

        /// <summary>
        /// 从网络共享中复制匹配模式的子文件夹内容到本地
        /// </summary>
        /// <param name="networkRoot">网络共享根路径，如 @"\\DATAPC03\RFAutoTestReport$\Chip Verification"</param>
        /// <param name="searchPattern">要匹配的一级子文件夹名称关键词（不区分大小写），如 "L004x"</param>
        /// <param name="localTargetBase">本地目标目录，如 @"C:\MyProject\CopiedReports"</param>
        /// <param name="username">可选：访问共享的用户名</param>
        /// <param name="password">可选：访问共享的密码</param>
        public void CopyMatchingSubFolders(
            string networkRoot,
            string searchPattern,
            string localTargetBase,
            string username = null,
            string password = null)
        {
            TargetPath = localTargetBase;
            if (string.IsNullOrWhiteSpace(networkRoot))
                throw new ArgumentException("网络路径不能为空", nameof(networkRoot));
            if (string.IsNullOrWhiteSpace(searchPattern))
                throw new ArgumentException("搜索关键词不能为空", nameof(searchPattern));
            if (string.IsNullOrWhiteSpace(localTargetBase))
                throw new ArgumentException("本地目标路径不能为空", nameof(localTargetBase));

            Directory.CreateDirectory(localTargetBase);
            ConnectToShare(networkRoot, username, password);

            try
            {
                Log($"?? 扫描网络目录: {networkRoot}");
                var topLevelDirs = Directory.GetDirectories(networkRoot, "*", SearchOption.TopDirectoryOnly);
                bool foundMatch = false;

                foreach (string dirPath in topLevelDirs)
                {
                    string dirName = Path.GetFileName(dirPath);
                    if (dirName.IndexOf(searchPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        foundMatch = true;
                        Log($"\n? 找到匹配的一级文件夹: {dirName}");
                        Log($"   路径: {dirPath}");

                        string[] subFolders;
                        try
                        {
                            subFolders = Directory.GetDirectories(dirPath, "*", SearchOption.TopDirectoryOnly);
                        }
                        catch (Exception ex)
                        {
                            Log($"   ? 无法读取子文件夹: {ex.Message}");
                            continue;
                        }

                        Log($"   ?? 发现 {subFolders.Length} 个子文件夹:");
                        foreach (var sf in subFolders)
                            Log($"     - {Path.GetFileName(sf)}");

                        if (subFolders.Length == 0)
                        {
                            Log("   ?? 无子文件夹可复制。");
                            continue;
                        }

                        foreach (string subFolder in subFolders)
                        {
                            string folderName = Path.GetFileName(subFolder);//得到问价夹名字， 而不是完整路径
                            if (folderName.IndexOf(searchPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                            {

                                string subFolderName = Path.GetFileName(subFolder);
                                string destPath = Path.Combine(localTargetBase, subFolderName);

                                try
                                {
                                    Log($"\n   ?? 正在复制: {subFolderName}");
                                    CopyDirectoryRecursive(subFolder, destPath);
                                    Log($"   ? 成功: {destPath}");
                                }
                                catch (Exception ex)
                                {
                                    Log($"   ? 复制失败 {subFolderName}: {ex.Message}");
                                }
                            }
                        }

                        // 通常只处理第一个匹配项（如 L004X），避免重复
                        break;
                    }
                }

                if (!foundMatch)
                {
                    Log($"? 未找到包含 \"{searchPattern}\" 的一级文件夹。");
                }
                else
                {
                    Log($"\n?? 所有子文件夹已复制到: {localTargetBase}");
                }
            }
            finally
            {
                DisconnectFromShare(networkRoot);
            }
        }

        private void ConnectToShare(string networkPath, string username, string password)
        {
            NETRESOURCE nr = new NETRESOURCE
            {
                dwType = 1, // RESOURCETYPE_DISK
                lpRemoteName = networkPath
            };

            int result = WNetAddConnection2(ref nr,
                string.IsNullOrEmpty(password) ? null : password,
                string.IsNullOrEmpty(username) ? null : username,
                0);

            if (result != 0)
            {
                throw new System.ComponentModel.Win32Exception(result,
                    $"无法连接到网络路径: {networkPath}");
            }
        }

        private void DisconnectFromShare(string networkPath)
        {
            try
            {
                WNetCancelConnection2(networkPath, 0, true);
            }
            catch
            {
                // 忽略断开失败（如连接本就不存在）
            }
        }

        private void CopyDirectoryRecursive(string sourceDir, string targetDir)
        {
            Directory.CreateDirectory(targetDir);

            //复制所有文件
            //foreach (string file in Directory.GetFiles(sourceDir))
            //{
            //    string fileName = Path.GetFileName(file);
            //    string destFile = Path.Combine(targetDir, fileName);
            //    File.Copy(file, destFile, overwrite: true);
            //}
            // 只复制文本文件
            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string extension = Path.GetExtension(file);
                if (TextFileExtensions.Contains(extension))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(targetDir, fileName);
                    File.Copy(file, destFile, overwrite: true);
                }
                // 否则跳过（不复制二进制文件如 .exe, .dll, .pdf, .zip 等）
            }

            //复制文件夹
            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                string subDirName = Path.GetFileName(subDir);
                string destSubDir = Path.Combine(targetDir, subDirName);
                CopyDirectoryRecursive(subDir, destSubDir);
            }
        }
        private static readonly HashSet<string> TextFileExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".txt",  ".csv", ".s2p", ".s3p", ".s4p", ".s5p"
               
                // 可根据实际需求增删
            };

    }
}
