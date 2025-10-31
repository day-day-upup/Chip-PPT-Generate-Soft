using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//  这个model 用于该软件的波形曲线的基础数据结构


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

        public string XLabel { set; get; }//X轴标签

        public string YLabel { set; get; }//Y轴标签


        //*****************X轴和Y轴的最大最小值*****************
        public double? xMin;

        public double? xMax;

        public double? yMin;

        public double? yMax;
        //*****************X轴和Y轴的最大最小值*****************


        public int? xAxisInterval;//x轴主坐标刻度间距

        public int? yAxisInterval;//y轴主坐标刻度间距

        public ScottPlot.Alignment Alignment { set; get; } = ScottPlot.Alignment.LowerRight;// 图列位置


        public void SetYAxisLimits(List<CurevelModel> Cureves)
        {
            if (Cureves == null || Cureves.Count == 0)
            {
                // 没有曲线数据，将边界设为 null
                xMin = xMax = yMin = yMax = null;
                return;
            }

            // 使用 LINQ SelectMany 将所有曲线的 XData 数组扁平化为一个序列，然后查找 Min/Max
            //xMin = Cureves
            //    .SelectMany(c => c.XData) // 扁平化所有 X 数组
            //    .Min();

            //xMax = Cureves
            //    .SelectMany(c => c.XData)
            //    .Max();

            // 对 YData 做同样的操作
            yMin = Cureves
                .SelectMany(c => c.YData)
                .Min();

            yMax = Cureves
                .SelectMany(c => c.YData)
                .Max();

            // 可以在这里加上一点余量（Padding），让图表更好看
            // xMin -= (xMax - xMin) * 0.05; 
            // yMax += (yMax - yMin) * 0.05;
        }

        /// <summary>
        /// 必须在对 Cureves 进行设置后调用，否则会导致边界计算错误
        /// </summary>
        public void SetYAxisLimits()
        {
            if (Cureves == null || Cureves.Count == 0)
            {
                // 没有曲线数据，将边界设为 null
                xMin = xMax = yMin = yMax = null;
                return;
            }

            // 使用 LINQ SelectMany 将所有曲线的 XData 数组扁平化为一个序列，然后查找 Min/Max
            //xMin = Cureves
            //    .SelectMany(c => c.XData) // 扁平化所有 X 数组
            //    .Min();

            //xMax = Cureves
            //    .SelectMany(c => c.XData)
            //    .Max();

            // 对 YData 做同样的操作
            yMin = Cureves
                .SelectMany(c => c.YData)
                .Min();

            yMax = Cureves
                .SelectMany(c => c.YData)
                .Max();

            // 可以在这里加上一点余量（Padding），让图表更好看
            //xMin -= (xMax - xMin) * 0.05;
            //yMax += (yMax - yMin) * 0.05;
        }



        public static bool CalculateFixedInterval(int min, int max, int targetDivisions, out int interval)
        {
            interval = 0;
            int range = max - min;

            // 范围必须大于 0 且目标份数必须大于 0
            if (range <= 0 || targetDivisions <= 0)
            {
                // 如果范围为 0 或负数，则间隔设为 1 或其他默认值
                interval = 1;
                return false;
            }

            // 1. 寻找能整除 range 的整数份数
            int actualDivisions = targetDivisions;
            bool found = false;

            // 从目标份数开始，向上下搜索一个能让 range 被整除的整数份数
            // 搜索范围：targetDivisions - 5 到 targetDivisions + 5
            for (int i = 0; i <= 5; i++)
            {
                // 尝试向下调整份数
                int divDown = targetDivisions - i;
                if (divDown > 0 && range % divDown == 0)
                {
                    actualDivisions = divDown;
                    found = true;
                    break;
                }

                // 尝试向上调整份数 (i=0 时已检查，这里 i > 0 时才检查)
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
                // 2. 计算最终的整数刻度间隔
                interval = range / actualDivisions;
                // 确保间隔至少为 1
                if (interval <= 0) interval = 1;
                return true;
            }
            else
            {
                // 如果在搜索范围内找不到合适的整数份数，则使用初始计算的四舍五入值
                // 这可能会导致边界不精确对齐，但能满足近似 10 份的要求
                interval = (int)Math.Round((double)range / targetDivisions);
                if (interval <= 0) interval = 1;
                return false;
            }
        }

        /// <summary>
        /// 根据原始 Min/Max 范围，计算整数刻度间隔，并调整 Min/Max 边界，
        /// 使新的边界能被间隔整除。
        /// 适用于 Min/Max 可以变动的场景（例如 Y 轴）。
        /// </summary>
        /// <param name="min">轴的原始最小值</param>
        /// <param name="max">轴的原始最大值</param>
        /// <param name="targetDivisions">目标刻度份数（例如 10）</param>
        /// <param name="newMin">输出：调整后的新最小值</param>
        /// <param name="newMax">输出：调整后的新最大值</param>
        /// <param name="interval">输出：计算得到的整数刻度间隔</param>
        public static void CalculateAdjustedRange(double min, double max, int targetDivisions,
                                                 out double newMin, out double newMax, out int interval)
        {
            double range = max - min;

            // 1. 计算初始的整数刻度间隔（四舍五入到整数）
            // 确保 range/targetDivisions > 0
            if (range <= 0 || targetDivisions <= 0)
            {
                // 边界情况处理
                interval = 1;
                newMin = Math.Floor(min);
                newMax = Math.Ceiling(max);
                if (newMax <= newMin) newMax = newMin + 1;
                return;
            }

            // 使用 Math.Round 获得一个接近的目标整数间隔
            interval = (int)Math.Round(range / targetDivisions);
            // 确保间隔至少为 1
            if (interval <= 0) interval = 1;

            // 2. 调整边界 Min' = floor(Min / Interval) * Interval
            newMin = Math.Floor(min / interval) * interval;

            // 3. 调整边界 Max' = ceil(Max / Interval) * Interval
            newMax = Math.Ceiling(max / interval) * interval;

            // 确保新范围至少包含一个刻度间隔
            if (newMax <= newMin)
            {
                newMax = newMin + interval;
            }
        }



        /// <summary>
        /// 寻找一个 "优美" 的刻度间隔（Nice Interval），它必须是 1, 2, 5, 10, 20, 50, 100... 等形式。
        /// 这个间隔将决定最终 yMin' 和 yMax' 都是 2 或 5 的倍数。
        /// </summary>
        /// <param name="range">数据范围 (yMax - yMin)</param>
        /// <param name="targetTicks">期望的刻度数量 (例如 10)</param>
        /// <returns>计算得到的优美整数刻度间隔</returns>
        private static int FindNiceInterval(double range, int targetTicks)
        {
            if (range <= 0 || targetTicks <= 0) return 1;

            // 1. 计算理想间隔
            double idealInterval = range / targetTicks;

            // 2. 找到理想间隔的指数（例如 12.3 -> 10, 0.45 -> 0.1, 1234 -> 1000）
            double exponent = Math.Floor(Math.Log10(idealInterval));
            double magnitude = Math.Pow(10, exponent); // 数量级

            // 3. 找到适合理想间隔的“优美分数”（1, 2, 或 5）
            double fractional = idealInterval / magnitude;

            int niceFraction;
            if (fractional <= 1.5)
            {
                niceFraction = 1; // 刻度间隔为 1 * magnitude
            }
            else if (fractional <= 3.0)
            {
                niceFraction = 2; // 刻度间隔为 2 * magnitude
            }
            else if (fractional <= 7.5)
            {
                niceFraction = 5; // 刻度间隔为 5 * magnitude
            }
            else
            {
                niceFraction = 10; // 刻度间隔为 10 * magnitude (即 1 * 10^(exponent+1))
            }

            // 4. 计算最终的优美刻度间隔
            int niceInterval = (int)Math.Round(niceFraction * magnitude);

            // 确保间隔是正整数
            return Math.Max(1, niceInterval);
        }

        /// <summary>
        /// 根据曲线数据计算 Y 轴刻度，并调整 Min/Max 边界，使其成为 Nice Interval 的整数倍。
        /// </summary>
        /// <param name="yMin">原始最小值</param>
        /// <param name="yMax">原始最大值</param>
        /// <param name="targetDivisions">目标刻度份数（例如 10）</param>
        /// <param name="newMin">输出：调整后的新最小值</param>
        /// <param name="newMax">输出：调整后的新最大值</param>
        /// <param name="interval">输出：计算得到的优美整数刻度间隔</param>
        public static void CalculateNiceRange(double yMin, double yMax, int targetDivisions,
                                              out double newMin, out double newMax, out int interval)
        {
            double range = yMax - yMin;

            // 1. 确定优美刻度间隔 (Interval)
            interval = FindNiceInterval(range, targetDivisions);

            // 2. 调整边界 Min' = floor(Min / Interval) * Interval
            // 这里的 Math.Floor 保证了 newMin 是 interval 的整数倍，且 <= yMin
            newMin = Math.Floor(yMin / interval) * interval;

            // 3. 调整边界 Max' = ceil(Max / Interval) * Interval
            // 这里的 Math.Ceiling 保证了 newMax 是 interval 的整数倍，且 >= yMax
            newMax = Math.Ceiling(yMax / interval) * interval;

            // 确保新范围至少包含一个刻度间隔
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

   

