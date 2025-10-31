using ScottPlot.Finance;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChipManualGenerationSogt
{
    internal class Common
    {
    }
    public class AmpfilierFilesbyGroup
    {
        //---------Temp: �����¶ȷ��� 
        //---------VD: ���յ�����������
        //----------��һ�������еĸ����ļ�
        public List<string> DataSparaFilePaths { get; set; } = new List<string>();
        public Dictionary<string, List<string>> DataSparabyTemp = new Dictionary<string, List<string>>();
        public Dictionary<string, List<string>> DataSparabyVD = new Dictionary<string, List<string>>();
       
 
        public List<string> NFFiles { get; set; } = new List<string>();
        public Dictionary<string, List<string>> NFbyTemp = new Dictionary<string, List<string>>();
        public Dictionary<string, List<string>> NFbyVD = new Dictionary<string, List<string>>();
        public List<string> PsatFiles { get; set; } = new List<string>();
        public Dictionary<string, List<string>> PsatbyTemp = new Dictionary<string, List<string>>();
        public Dictionary<string, List<string>> PsatbyVD = new Dictionary<string, List<string>>();
        public List<string> OIP3Files { get; set; } = new List<string>();
        public Dictionary<string, List<string>> OIP3byTemp = new Dictionary<string, List<string>>();
        public Dictionary<string, List<string>> OIP3byVD = new Dictionary<string, List<string>>();
        public List<string> PxdBFiles { get; set; } = new List<string>();
        public Dictionary<string, List<string>> PxdBbyTemp = new Dictionary<string, List<string>>();
        public Dictionary<string, List<string>> PxdBbyVD = new Dictionary<string, List<string>>();
    }


    public class DataConverter
    {
        /// <summary>
        /// �� List<(string name, string info)> �ṹת��Ϊ string[,] (��ά����)��
        /// ����ṹ��[��, 2]
        ///           [i, 0] = name
        ///           [i, 1] = info
        /// </summary>
        /// <param name="data">Ҫת����Ԫ���б�</param>
        /// <returns>�����������ݵĶ�ά�ַ������顣</returns>
        static public string[,] ConvertListToTwoDArray(List<(string name, string info)> data)
        {
            // 1. ��������Ƿ�Ϊ��
            if (data == null || data.Count == 0)
            {
                // ����б�Ϊ�գ�����һ�� 0 �� 0 �еĿ�����
                return new string[0, 0];
            }

            // 2. ȷ��ά��
            int rows = data.Count;
            const int cols = 2; // �̶�Ϊ���У�name �� info

            // 3. ����Ŀ���ά����
            string[,] resultArray = new string[rows, cols];

            // 4. �����б���ֵ������
            for (int i = 0; i < rows; i++)
            {
                var item = data[i];

                // ��һ�У�name
                resultArray[i, 0] = item.name;

                // �ڶ��У�info
                resultArray[i, 1] = item.info;
            }

            return resultArray;
        }


        /// <summary>
        /// �� List<(string vd, string vg, string idd)> �ṹת��Ϊ string[,] (��ά����)��
        /// ����ṹ��[��, 3]
        ///           [i, 0] = vd
        ///           [i, 1] = vg
        ///           [i, 2] = idd
        /// </summary>
        /// <param name="data">Ҫת����Ԫ���б�</param>
        /// <returns>�����������ݵĶ�ά�ַ������顣</returns>
        static public string[,] ConvertThreeElementListToTwoDArray(List<(string vd, string vg, string idd)> data)
        {
            // 1. ��������Ƿ�Ϊ��
            if (data == null || data.Count == 0)
            {
                // ����б�Ϊ�գ�����һ�� 0 �� 0 �еĿ�����
                return new string[0, 0];
            }

            // 2. ȷ��ά��
            int rows = data.Count;
            const int cols = 3; // �̶�Ϊ���У�vd, vg, idd

            // 3. ����Ŀ���ά����
            string[,] resultArray = new string[rows, cols];

            // 4. �����б���ֵ������
            for (int i = 0; i < rows; i++)
            {
                var item = data[i];

                // �� 0 �У�vd
                resultArray[i, 0] = item.vd;

                // �� 1 �У�vg
                resultArray[i, 1] = item.vg;

                // �� 2 �У�idd
                resultArray[i, 2] = item.idd;
            }

            return resultArray;
        }
    }

    public class LogModel
    {
        public DateTime TimeStamp { get; set; }
        public string UserName { get; set; }

        public string TaskId { get; set; }
        public string TaskName { get; set; }

        public string Level { get; set; }
        public string Message { get; set; }

        public string PN { get; set; }

        public string SN { get; set; }

        public string ChipNumber { get; set; }

    }

    public static class LogLevels
    {


        /// <summary>������Ϣ������������ϡ�</summary>
        public const string Debug = "DEBUG";

        /// <summary>����ҵ��������Ϣ��</summary>
        public const string Info = "INFO";

        /// <summary>Ǳ�����⣬��Ӧ�ÿ��Լ������С�</summary>
        public const string Warning = "WARNING";

        /// <summary>������Ҫ����</summary>
        public const string Error = "ERROR";

        /// <summary>���ش��󣬵���Ӧ�ó������Ĺ��ܱ�����</summary>
        public const string Fatal = "FATAL";
    }

    public enum UserPriority
    {
        Admin = 0,
        DataProvider = 1,
        PptMaker = 2,
        Reviewer = 3
    }

    /// <summary>
    /// ����ǲ���ģ�ͣ� ������¼���������е�һЩ��Ϣ�� ����������� ���ɵ�ppt ·���ȵ�
    /// </summary>
    public class OperationModel
    { 
        public int TaskID { get; set; }
        public string TaskName { get; set; }

        public DateTime TimeStamp { get; set; }

        public DateTime? StartDateTime { get; set; }

        public DateTime? EndDateTime { get; set; }
        public string PN { get; set; }
        public string SN { get; set; }

        public bool DataReady { get; set; }

        public string Condition { get; set; }

        public string Data { get; set; }
        public string SourceFiles { get; set; }// ��¼���ڼ�¼���ݵ�ԭʼ����

        public  bool FileReady { get; set; }

        public string PptPath { get; set; }

    }

}


