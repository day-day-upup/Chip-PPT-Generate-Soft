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

        // ��־���ί�У���ѡ��
        public Action<string> Log { get; set; } = msg => Console.WriteLine(msg);

        /// <summary>
        /// �����繲���и���ƥ��ģʽ�����ļ������ݵ�����
        /// </summary>
        /// <param name="networkRoot">���繲���·������ @"\\DATAPC03\RFAutoTestReport$\Chip Verification"</param>
        /// <param name="searchPattern">Ҫƥ���һ�����ļ������ƹؼ��ʣ������ִ�Сд������ "L004x"</param>
        /// <param name="localTargetBase">����Ŀ��Ŀ¼���� @"C:\MyProject\CopiedReports"</param>
        /// <param name="username">��ѡ�����ʹ�����û���</param>
        /// <param name="password">��ѡ�����ʹ��������</param>
        public void CopyMatchingSubFolders(
            string networkRoot,
            string searchPattern,
            string localTargetBase,
            string username = null,
            string password = null)
        {
            TargetPath = localTargetBase;
            if (string.IsNullOrWhiteSpace(networkRoot))
                throw new ArgumentException("����·������Ϊ��", nameof(networkRoot));
            if (string.IsNullOrWhiteSpace(searchPattern))
                throw new ArgumentException("�����ؼ��ʲ���Ϊ��", nameof(searchPattern));
            if (string.IsNullOrWhiteSpace(localTargetBase))
                throw new ArgumentException("����Ŀ��·������Ϊ��", nameof(localTargetBase));

            Directory.CreateDirectory(localTargetBase);
            ConnectToShare(networkRoot, username, password);

            try
            {
                Log($"?? ɨ������Ŀ¼: {networkRoot}");
                var topLevelDirs = Directory.GetDirectories(networkRoot, "*", SearchOption.TopDirectoryOnly);
                bool foundMatch = false;

                foreach (string dirPath in topLevelDirs)
                {
                    string dirName = Path.GetFileName(dirPath);
                    if (dirName.IndexOf(searchPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        foundMatch = true;
                        Log($"\n? �ҵ�ƥ���һ���ļ���: {dirName}");
                        Log($"   ·��: {dirPath}");

                        string[] subFolders;
                        try
                        {
                            subFolders = Directory.GetDirectories(dirPath, "*", SearchOption.TopDirectoryOnly);
                        }
                        catch (Exception ex)
                        {
                            Log($"   ? �޷���ȡ���ļ���: {ex.Message}");
                            continue;
                        }

                        Log($"   ?? ���� {subFolders.Length} �����ļ���:");
                        foreach (var sf in subFolders)
                            Log($"     - {Path.GetFileName(sf)}");

                        if (subFolders.Length == 0)
                        {
                            Log("   ?? �����ļ��пɸ��ơ�");
                            continue;
                        }

                        foreach (string subFolder in subFolders)
                        {
                            string folderName = Path.GetFileName(subFolder);//�õ��ʼۼ����֣� ����������·��
                            if (folderName.IndexOf(searchPattern, StringComparison.OrdinalIgnoreCase) >= 0)
                            {

                                string subFolderName = Path.GetFileName(subFolder);
                                string destPath = Path.Combine(localTargetBase, subFolderName);

                                try
                                {
                                    Log($"\n   ?? ���ڸ���: {subFolderName}");
                                    CopyDirectoryRecursive(subFolder, destPath);
                                    Log($"   ? �ɹ�: {destPath}");
                                }
                                catch (Exception ex)
                                {
                                    Log($"   ? ����ʧ�� {subFolderName}: {ex.Message}");
                                }
                            }
                        }

                        // ͨ��ֻ�����һ��ƥ����� L004X���������ظ�
                        break;
                    }
                }

                if (!foundMatch)
                {
                    Log($"? δ�ҵ����� \"{searchPattern}\" ��һ���ļ��С�");
                }
                else
                {
                    Log($"\n?? �������ļ����Ѹ��Ƶ�: {localTargetBase}");
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
                    $"�޷����ӵ�����·��: {networkPath}");
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
                // ���ԶϿ�ʧ�ܣ������ӱ��Ͳ����ڣ�
            }
        }

        private void CopyDirectoryRecursive(string sourceDir, string targetDir)
        {
            Directory.CreateDirectory(targetDir);

            //���������ļ�
            //foreach (string file in Directory.GetFiles(sourceDir))
            //{
            //    string fileName = Path.GetFileName(file);
            //    string destFile = Path.Combine(targetDir, fileName);
            //    File.Copy(file, destFile, overwrite: true);
            //}
            // ֻ�����ı��ļ�
            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string extension = Path.GetExtension(file);
                if (TextFileExtensions.Contains(extension))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(targetDir, fileName);
                    File.Copy(file, destFile, overwrite: true);
                }
                // ���������������ƶ������ļ��� .exe, .dll, .pdf, .zip �ȣ�
            }

            //�����ļ���
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
               
                // �ɸ���ʵ��������ɾ
            };

    }
}
