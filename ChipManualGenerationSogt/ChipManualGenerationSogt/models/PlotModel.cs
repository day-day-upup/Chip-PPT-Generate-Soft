using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//  ���model ���ڸ�����Ĳ������ߵĻ������ݽṹ


namespace ChipManualGenerationSogt.models
{
    public class CurevelModel
    { 
        public string Legend { set; get; }
        public double[] XData { set; get; }
        public double[] YData { set; get; }
    
    }

    public class PlotModel
    {
        public List<CurevelModel> Cureves { set; get; } = new List<CurevelModel>();
        
        public string Title { set; get; }

        public string XLabel { set; get; }//X���ǩ

        public string YLabel { set; get; }//Y���ǩ


        //*****************X���Y��������Сֵ*****************
        public double? xMin;

        public double? xMax;

        public double? yMin;

        public double? yMax;
        //*****************X���Y��������Сֵ*****************


        public int? xAxisInterval;//x��������̶ȼ��

        public int? yAxisInterval;//y��������̶ȼ��

        public ScottPlot.Alignment Alignment { set; get; } = ScottPlot.Alignment.LowerRight;// ͼ��λ��


        public void SetYAxisLimits(List<CurevelModel> Cureves)
        {
            if (Cureves == null || Cureves.Count == 0)
            {
                // û���������ݣ����߽���Ϊ null
                xMin = xMax = yMin = yMax = null;
                return;
            }

            // ʹ�� LINQ SelectMany ���������ߵ� XData �����ƽ��Ϊһ�����У�Ȼ����� Min/Max
            //xMin = Cureves
            //    .SelectMany(c => c.XData) // ��ƽ������ X ����
            //    .Min();

            //xMax = Cureves
            //    .SelectMany(c => c.XData)
            //    .Max();

            // �� YData ��ͬ���Ĳ���
            yMin = Cureves
                .SelectMany(c => c.YData)
                .Min();

            yMax = Cureves
                .SelectMany(c => c.YData)
                .Max();

            // �������������һ��������Padding������ͼ����ÿ�
            // xMin -= (xMax - xMin) * 0.05; 
            // yMax += (yMax - yMin) * 0.05;
        }

        /// <summary>
        /// �����ڶ� Cureves �������ú���ã�����ᵼ�±߽�������
        /// </summary>
        public void SetYAxisLimits()
        {
            if (Cureves == null || Cureves.Count == 0)
            {
                // û���������ݣ����߽���Ϊ null
                xMin = xMax = yMin = yMax = null;
                return;
            }

            // ʹ�� LINQ SelectMany ���������ߵ� XData �����ƽ��Ϊһ�����У�Ȼ����� Min/Max
            //xMin = Cureves
            //    .SelectMany(c => c.XData) // ��ƽ������ X ����
            //    .Min();

            //xMax = Cureves
            //    .SelectMany(c => c.XData)
            //    .Max();

            // �� YData ��ͬ���Ĳ���
            yMin = Cureves
                .SelectMany(c => c.YData)
                .Min();

            yMax = Cureves
                .SelectMany(c => c.YData)
                .Max();

            // �������������һ��������Padding������ͼ����ÿ�
            //xMin -= (xMax - xMin) * 0.05;
            //yMax += (yMax - yMin) * 0.05;
        }



        public static bool CalculateFixedInterval(int min, int max, int targetDivisions, out int interval)
        {
            interval = 0;
            int range = max - min;

            // ��Χ������� 0 ��Ŀ������������ 0
            if (range <= 0 || targetDivisions <= 0)
            {
                // �����ΧΪ 0 ������������Ϊ 1 ������Ĭ��ֵ
                interval = 1;
                return false;
            }

            // 1. Ѱ�������� range ����������
            int actualDivisions = targetDivisions;
            bool found = false;

            // ��Ŀ�������ʼ������������һ������ range ����������������
            // ������Χ��targetDivisions - 5 �� targetDivisions + 5
            for (int i = 0; i <= 5; i++)
            {
                // �������µ�������
                int divDown = targetDivisions - i;
                if (divDown > 0 && range % divDown == 0)
                {
                    actualDivisions = divDown;
                    found = true;
                    break;
                }

                // �������ϵ������� (i=0 ʱ�Ѽ�飬���� i > 0 ʱ�ż��)
                if (i > 0)
                {
                    int divUp = targetDivisions + i;
                    if (range % divUp == 0)
                    {
                        actualDivisions = divUp;
                        found = true;
                        break;
                    }
                }
            }

            if (found)
            {
                // 2. �������յ������̶ȼ��
                interval = range / actualDivisions;
                // ȷ���������Ϊ 1
                if (interval <= 0) interval = 1;
                return true;
            }
            else
            {
                // �����������Χ���Ҳ������ʵ�������������ʹ�ó�ʼ�������������ֵ
                // ����ܻᵼ�±߽粻��ȷ���룬����������� 10 �ݵ�Ҫ��
                interval = (int)Math.Round((double)range / targetDivisions);
                if (interval <= 0) interval = 1;
                return false;
            }
        }

        /// <summary>
        /// ����ԭʼ Min/Max ��Χ�����������̶ȼ���������� Min/Max �߽磬
        /// ʹ�µı߽��ܱ����������
        /// ������ Min/Max ���Ա䶯�ĳ��������� Y �ᣩ��
        /// </summary>
        /// <param name="min">���ԭʼ��Сֵ</param>
        /// <param name="max">���ԭʼ���ֵ</param>
        /// <param name="targetDivisions">Ŀ��̶ȷ��������� 10��</param>
        /// <param name="newMin">����������������Сֵ</param>
        /// <param name="newMax">�����������������ֵ</param>
        /// <param name="interval">���������õ��������̶ȼ��</param>
        public static void CalculateAdjustedRange(double min, double max, int targetDivisions,
                                                 out double newMin, out double newMax, out int interval)
        {
            double range = max - min;

            // 1. �����ʼ�������̶ȼ�����������뵽������
            // ȷ�� range/targetDivisions > 0
            if (range <= 0 || targetDivisions <= 0)
            {
                // �߽��������
                interval = 1;
                newMin = Math.Floor(min);
                newMax = Math.Ceiling(max);
                if (newMax <= newMin) newMax = newMin + 1;
                return;
            }

            // ʹ�� Math.Round ���һ���ӽ���Ŀ���������
            interval = (int)Math.Round(range / targetDivisions);
            // ȷ���������Ϊ 1
            if (interval <= 0) interval = 1;

            // 2. �����߽� Min' = floor(Min / Interval) * Interval
            newMin = Math.Floor(min / interval) * interval;

            // 3. �����߽� Max' = ceil(Max / Interval) * Interval
            newMax = Math.Ceiling(max / interval) * interval;

            // ȷ���·�Χ���ٰ���һ���̶ȼ��
            if (newMax <= newMin)
            {
                newMax = newMin + interval;
            }
        }



        /// <summary>
        /// Ѱ��һ�� "����" �Ŀ̶ȼ����Nice Interval������������ 1, 2, 5, 10, 20, 50, 100... ����ʽ��
        /// ���������������� yMin' �� yMax' ���� 2 �� 5 �ı�����
        /// </summary>
        /// <param name="range">���ݷ�Χ (yMax - yMin)</param>
        /// <param name="targetTicks">�����Ŀ̶����� (���� 10)</param>
        /// <returns>����õ������������̶ȼ��</returns>
        private static int FindNiceInterval(double range, int targetTicks)
        {
            if (range <= 0 || targetTicks <= 0) return 1;

            // 1. ����������
            double idealInterval = range / targetTicks;

            // 2. �ҵ���������ָ�������� 12.3 -> 10, 0.45 -> 0.1, 1234 -> 1000��
            double exponent = Math.Floor(Math.Log10(idealInterval));
            double magnitude = Math.Pow(10, exponent); // ������

            // 3. �ҵ��ʺ��������ġ�������������1, 2, �� 5��
            double fractional = idealInterval / magnitude;

            int niceFraction;
            if (fractional <= 1.5)
            {
                niceFraction = 1; // �̶ȼ��Ϊ 1 * magnitude
            }
            else if (fractional <= 3.0)
            {
                niceFraction = 2; // �̶ȼ��Ϊ 2 * magnitude
            }
            else if (fractional <= 7.5)
            {
                niceFraction = 5; // �̶ȼ��Ϊ 5 * magnitude
            }
            else
            {
                niceFraction = 10; // �̶ȼ��Ϊ 10 * magnitude (�� 1 * 10^(exponent+1))
            }

            // 4. �������յ������̶ȼ��
            int niceInterval = (int)Math.Round(niceFraction * magnitude);

            // ȷ�������������
            return Math.Max(1, niceInterval);
        }

        /// <summary>
        /// �����������ݼ��� Y ��̶ȣ������� Min/Max �߽磬ʹ���Ϊ Nice Interval ����������
        /// </summary>
        /// <param name="yMin">ԭʼ��Сֵ</param>
        /// <param name="yMax">ԭʼ���ֵ</param>
        /// <param name="targetDivisions">Ŀ��̶ȷ��������� 10��</param>
        /// <param name="newMin">����������������Сֵ</param>
        /// <param name="newMax">�����������������ֵ</param>
        /// <param name="interval">���������õ������������̶ȼ��</param>
        public static void CalculateNiceRange(double yMin, double yMax, int targetDivisions,
                                              out double newMin, out double newMax, out int interval)
        {
            double range = yMax - yMin;

            // 1. ȷ�������̶ȼ�� (Interval)
            interval = FindNiceInterval(range, targetDivisions);

            // 2. �����߽� Min' = floor(Min / Interval) * Interval
            // ����� Math.Floor ��֤�� newMin �� interval ������������ <= yMin
            newMin = Math.Floor(yMin / interval) * interval;

            // 3. �����߽� Max' = ceil(Max / Interval) * Interval
            // ����� Math.Ceiling ��֤�� newMax �� interval ������������ >= yMax
            newMax = Math.Ceiling(yMax / interval) * interval;

            // ȷ���·�Χ���ٰ���һ���̶ȼ��
            if (newMax <= newMin)
            {
                newMax = newMin + interval;
            }
        }
    }

    public enum LegendTextType
    { 
        Temp,
        Elec
    }
}

   

