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
    //**************************  S2P �ļ������� **********
    public class S_Parameters
    {
        public double FreqGHz { get; set; }   // Ƶ�ʣ�GHz��
        public double DbValue { get; set; }   // ���ȣ�dB��
        public double PhaseValue { get; set; } // ��λ���ȣ�
    }

    // ������ͨ����pxdb�� psat ��nf�ȣ� ���Ĳ���һ�����ݣ� �������ʾ
    public class Pair_Parameters
    {
        public double FreqGHz { get; set; }   // Ƶ�ʣ�GHz��
        public double Value { get; set; }   // ���ȣ�dB��
    }


    //*************************************  Txt File ������ **********
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
                        trimmedLine.StartsWith("!") ||
                        trimmedLine.StartsWith("#"))
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
            catch (Exception ex) 
            {
                // �ɼ�¼��־��Logger?.LogError(ex, "��ȡ S2P �ļ�ʧ��");
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
                throw new FileNotFoundException($"�ļ�δ�ҵ�: {filePath}");

            points.Clear(); // ��վ�����

            var lines = File.ReadAllLines(filePath);
            const int dataStartLine = 4; // ��4�У�����Ϊ3��
            
            for (int i = dataStartLine; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrEmpty(line))
                    continue; // ��������

                // ���հ��ַ����ո��Ʊ���ȣ��ָ�
                string[] parts = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length < 2)
                    throw new FormatException($"�����и�ʽ���󣨵� {i + 1} �У�: '{line}'");

                // ʹ�ò����Ļ���InvariantCulture��������������������������Ӱ�죨�綺�� vs �㣩
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
                    throw new FormatException($"�޷�������ֵ���� {i + 1} �У�: '{line}'");
                }
            }
        }

    }
}
