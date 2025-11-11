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
        //---------Temp: 按照温度分类 
        //---------VD: 按照电气参数分类
        //----------第一个是所有的该类文件
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
        /// 将 List<(string name, string info)> 结构转换为 string[,] (二维数组)。
        /// 数组结构：[行, 2]
        ///           [i, 0] = name
        ///           [i, 1] = info
        /// </summary>
        /// <param name="data">要转换的元组列表。</param>
        /// <returns>包含两列数据的二维字符串数组。</returns>
        static public string[,] ConvertListToTwoDArray(List<(string name, string info)> data)
        {
            // 1. 检查输入是否为空
            if (data == null || data.Count == 0)
            {
                // 如果列表为空，返回一个 0 行 0 列的空数组
                return new string[0, 0];
            }

            // 2. 确定维度
            int rows = data.Count;
            const int cols = 2; // 固定为两列：name 和 info

            // 3. 创建目标二维数组
            string[,] resultArray = new string[rows, cols];

            // 4. 遍历列表并赋值给数组
            for (int i = 0; i < rows; i++)
            {
                var item = data[i];

                // 第一列：name
                resultArray[i, 0] = item.name;

                // 第二列：info
                resultArray[i, 1] = item.info;
            }

            return resultArray;
        }


        /// <summary>
        /// 将 List<(string vd, string vg, string idd)> 结构转换为 string[,] (二维数组)。
        /// 数组结构：[行, 3]
        ///           [i, 0] = vd
        ///           [i, 1] = vg
        ///           [i, 2] = idd
        /// </summary>
        /// <param name="data">要转换的元组列表。</param>
        /// <returns>包含三列数据的二维字符串数组。</returns>
        static public string[,] ConvertThreeElementListToTwoDArray(List<(string vd, string vg, string idd)> data)
        {
            // 1. 检查输入是否为空
            if (data == null || data.Count == 0)
            {
                // 如果列表为空，返回一个 0 行 0 列的空数组
                return new string[0, 0];
            }

            // 2. 确定维度
            int rows = data.Count;
            const int cols = 3; // 固定为三列：vd, vg, idd

            // 3. 创建目标二维数组
            string[,] resultArray = new string[rows, cols];

            // 4. 遍历列表并赋值给数组
            for (int i = 0; i < rows; i++)
            {
                var item = data[i];

                // 第 0 列：vd
                resultArray[i, 0] = item.vd;

                // 第 1 列：vg
                resultArray[i, 1] = item.vg;

                // 第 2 列：idd
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


        /// <summary>调试信息，用于问题诊断。</summary>
        public const string Debug = "DEBUG";

        /// <summary>常规业务流程信息。</summary>
        public const string Info = "INFO";

        /// <summary>潜在问题，但应用可以继续运行。</summary>
        public const string Warning = "WARNING";

        /// <summary>错误，需要处理。</summary>
        public const string Error = "ERROR";

        /// <summary>严重错误，导致应用程序或核心功能崩溃。</summary>
        public const string Fatal = "FATAL";
    }

    public static class TaskStatus
    {
        public const string NotCommited = "Not Commited";
        public const string PPTReady = "PPT Ready For Generate";
        public const string Completed = "Completed";
       
    }   
    public enum UserPriority
    {
        Admin = 0,
        DataProvider = 1,
        PptMaker = 2,
        Reviewer = 3
    }

    /// <summary>
    /// 这个是操作模型， 用来记录操作过程中的一些信息， 如操作条件， 生成的ppt 路径等等
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
        public string SourceFiles { get; set; }// 记录用于记录数据的原始数据

        public  bool FileReady { get; set; }

        public string PptPath { get; set; }

    }



   

// 假设数据库中的表名为 'Tasks'
public class TaskSqlServerModel
    {
        // 标识符：自增主键
        public int ID { get; set; }

        // 任务信息
        public string PPTModel { get; set; }
        public string TaskName { get; set; }
        public string Status { get; set; }
        public string Level { get; set; }
        public string Major { get; set; }

        // 注意：根据你的描述，Minor 可能是字符串，代表子设备名
        public string Minor { get; set; }

        // 时间信息
        public DateTime StartDate { get; set; }
        public DateTime? EndDate { get; set; }

        // 状态布尔值
        public bool DataStatus { get; set; }
        public bool FilesStatus { get; set; }

        // 条件/配置
        public string Conditions { get; set; } // 注意拼写，如果数据库是 'Conditions'，此处也应修改
        public string PptName { get; set; }

        public bool TableUpdate { get; set; }
    }
}


