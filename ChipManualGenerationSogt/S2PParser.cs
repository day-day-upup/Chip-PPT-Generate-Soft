using System;
using System.Collections.Generic;
using System.IO;
using DT = System.Data;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace ChipManualGenerationSogt
{
    //**************************  S2P 文件解析器 **********
    public class S_Parameters
    {
        public double FreqGHz { get; set; }   // 频率（GHz）
        public double DbValue { get; set; }   // 幅度（dB）
        public double PhaseValue { get; set; } // 相位（度）
    }

    // 用于普通的如pxdb， psat ，nf等， 它的参是一对数据， 用这个表示
    public class Pair_Parameters
    {
        public double FreqGHz { get; set; }   // 频率（GHz）
        public double Value { get; set; }   // 幅度（dB）
    }


    //*************************************  Txt File 解析器 **********
    public class S2PParser
    {
        private List<S_Parameters> _s11;

        public List<S_Parameters> S11 { get { return _s11; } }

        private List<S_Parameters> _s21;

        public List<S_Parameters> S21 { get { return _s21; } }

        private List<S_Parameters> _s12;

        public List<S_Parameters> S12 { get { return _s12; } }


        private List<S_Parameters> _s22;

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
                        trimmedLine.StartsWith("!") ||
                        trimmedLine.StartsWith("#"))
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
            catch (Exception ex) 
            {
                // 可记录日志：Logger?.LogError(ex, "读取 S2P 文件失败");
                return false;
            }
        }


    }

    public class TextFileParser
    {

        private List<Pair_Parameters> points = new List<Pair_Parameters>();

        public List<Pair_Parameters> Points { get { return points; } }
        public void Parse(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"文件未找到: {filePath}");

            points.Clear(); // 清空旧数据

            var lines = File.ReadAllLines(filePath);
            const int dataStartLine = 4; // 第4行（索引为3）
            
            for (int i = dataStartLine; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrEmpty(line))
                    continue; // 跳过空行

                // 按空白字符（空格、制表符等）分割
                string[] parts = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length < 2)
                    throw new FormatException($"数据行格式错误（第 {i + 1} 行）: '{line}'");

                // 使用不变文化（InvariantCulture）解析浮点数，避免区域设置影响（如逗号 vs 点）
                if (double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double freq) &&
                    double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                {
                    points.Add(new Pair_Parameters
                    {
                        FreqGHz = freq ,
                        Value = value
                    });
                }
                
                else
                {
                    throw new FormatException($"无法解析数值（第 {i + 1} 行）: '{line}'");
                }
            }
        }

    }
}
