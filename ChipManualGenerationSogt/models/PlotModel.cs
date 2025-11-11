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




        public int? xAxisInterval;//x轴主坐标刻度间距

        public double? yAxisInterval;//y轴主坐标刻度间距

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

        public static void S11CalculateNiceRange(double yMin, double yMax, int targetDivisions,
                                        out double newMin, out double newMax, out double interval)
        {
            // 1. 固定最大值
            newMax = 0.0;

            // 确保 yMin 不会高于 newMax，并计算实际范围
            double effectiveMin = Math.Min(yMin, newMax);
            double effectiveRange = newMax - effectiveMin;

            // 2. 确定理想的刻度间隔 (Interval)
            // 目标是 10 个刻度，所以 interval 应该约等于 effectiveRange / 10
            double roughInterval = effectiveRange / (double)targetDivisions;

            // S11 刻度通常是 5, 10, 20... 的倍数。我们强制向上取整到最接近的 5 的倍数。
            int stepFactor = 5;
            interval = (int)(Math.Ceiling(roughInterval / (double)stepFactor) * stepFactor);

            // 确保 interval 至少为 5 (或一个合理的最小值)
            if (interval == 0)
            {
                interval = 5;
            }

            // 3. 计算强制扩大后的新最小值 (newMin)

            // 如果当前的有效范围 effectiveRange 较小，导致刻度数量不足 targetDivisions (10 个)
            // 我们需要强制扩大范围，直到范围能容纳 targetDivisions 数量的刻度。

            // 目标范围 = targetDivisions * interval
            double targetRange = (double)targetDivisions * interval;

            // 如果有效范围小于目标范围，我们需要扩展它。
            if (effectiveRange < targetRange)
            {
                // 目标 newMin = newMax - targetRange
                newMin = newMax - targetRange; // 例如: 0 - 50 = -50
            }
            else
            {
                // 如果有效范围足够大，我们只需要将 newMin 向下取整到 interval 的倍数
                // 例如: yMin = -58, interval = 5. newMin = floor(-58/5)*5 = -60
                newMin = Math.Floor(effectiveMin / interval) * interval;
            }

            // 4. 最终检查：确保 newMin < newMax
            if (newMax <= newMin)
            {
                // 只有当 yMin 意外为 0 且 interval=5 时发生
                newMin = newMax - interval; // newMin = -5
            }

            //newMax = 0.0;

            //double subMax = Math.Ceiling(yMax);
            //double subMin = Math.Floor(yMin);
            //double range = subMax - subMin;

            //interval = (int)Math.Ceiling((decimal)(range / 4));

            //while (Math.Abs(subMax) % interval != 0)
            //{

            //    subMax++;
            //}

            //newMin = subMax - 10 * interval + (newMax - subMax);

        }


        public static void S21CalculateNiceRange(double yMin, double yMax, int targetDivisions,
                                        out double newMin, out double newMax, out double interval)
        {
            //// 1. 固定最小值
            //newMin = 0.0;

            //// 确保 yMax 不会低于 newMin，并计算实际范围
            //double effectiveMax = Math.Max(yMax, newMin);
            //double effectiveRange = effectiveMax - newMin;

            //// 如果范围极小，设置一个默认范围以避免计算错误或 interval=0
            //if (effectiveRange <= 1e-9)
            //{
            //    effectiveRange = 4.0; // 默认范围 4 dB，方便计算 1, 2, 4 的间隔
            //}

            //// 2. 确定理想的刻度间隔 (Interval)
            //// 目标间隔的粗略估计： range / 10
            //double roughInterval = effectiveRange / (double)targetDivisions;

            //// a) 确定刻度间隔的科学计数法指数 (Power)
            //// 10^power < roughInterval <= 10^(power+1)
            //double power = Math.Floor(Math.Log10(roughInterval));
            //double exponent = Math.Pow(10, power); // 例如: 10, 1, 0.1, 0.01...

            //// b) 确定刻度间隔的前导数字 (Mantissa)
            //double mantissa = roughInterval / exponent; // 介于 1 到 10 之间

            //// c) 优美数 (Nice Numbers) 必须是 1, 2, 4, 8 (或 10)
            //double niceMantissa;
            //if (mantissa < 1.0) niceMantissa = 1.0;
            //else if (mantissa <= 2.0) niceMantissa = 2.0;
            //else if (mantissa <= 4.0) niceMantissa = 4.0;
            //else if (mantissa <= 8.0) niceMantissa = 8.0; // 8 也是 2 的倍数，可以作为备选
            //else niceMantissa = 10.0; // 如果大于 8，则取 10

            //// d) 计算最终的间隔，并转换为 int
            //double niceInterval = niceMantissa * exponent;
            //interval = Math.Max(1, (int)Math.Round(niceInterval));

            //// 确保 interval 至少为 1
            //if (interval == 0)
            //{
            //    interval = 1;
            //}

            //// 3. 计算强制扩大后的新最大值 (newMax) 

            //// 目标范围 = targetDivisions * interval
            //double targetRange = (double)targetDivisions * interval;

            //// 如果有效范围小于目标范围，我们需要扩展它。
            //if (effectiveRange < targetRange)
            //{
            //    // 目标 newMax = newMin + targetRange
            //    newMax = newMin + targetRange;
            //}
            //else
            //{
            //    // 如果有效范围足够大，我们只需要将 yMax 向上取整到 interval 的倍数
            //    newMax = Math.Ceiling(effectiveMax / interval) * interval;
            //}

            //// 4. 最终检查：确保 newMin < newMax
            //if (newMax <= newMin)
            //{
            //    newMax = newMin + interval;
            //}


           

            double subMax = Math.Ceiling(yMax);
            interval = (int)Math.Ceiling((decimal)(subMax / 9));
            newMin = 0.0;
           
            newMax = 10 * interval;
           
        }
        /// <summary>
        /// [ymin -x] >= 2(ymax-ymin) + 10x   抽象函数表达式 式这个 x为 interval
        /// </summary>
        /// <param name="yMin"></param>
        /// <param name="yMax"></param>
        /// <param name="targetDivisions"></param>
        /// <param name="newMin"></param>
        /// <param name="newMax"></param>
        /// <param name="interval"></param>
        public static void PxdbCalculateNiceRange(double yMin, double yMax, int targetDivisions,
                                        out double newMin, out double newMax, out double interval)
        {

            // 1. 计算数据实际范围和粗略间隔
            //double range = yMax - yMin;
            //int intRange = (int)Math.Ceiling(range);
            //double maxX = InequalitySolver.SolveForMaxX(yMin, yMax);
            //interval = (int)Math.Ceiling(Math.Abs(maxX));

            //newMin = Math.Floor(yMin - interval)-3*interval;
            //newMax = newMin +10 * interval;

            double subMax = Math.Ceiling(yMax);
            double subMin = Math.Floor(yMin);
            double range = subMax - subMin;

            interval = (int)Math.Ceiling((decimal)(range / 4));

            while (subMax % interval != 0)
            {

                subMax++;
            }

            newMin = subMax - 10 * interval;
            newMax = subMax;

        }

        public static void PsatCalculateNiceRange(double yMin, double yMax, int targetDivisions,
                                            out double newMin, out double newMax, out double interval)
        {
            double subMax = Math.Ceiling(yMax);
            double subMin = Math.Floor(yMin);
            double range = subMax - subMin;
           
            interval = (int)Math.Ceiling((decimal)(range /4));

            while (subMax % interval != 0)
            { 
            
                subMax++;
            }

            newMin = subMax -10*interval;
            newMax = subMax;

        }

        public static int? newInterval(double yMin, double yMax)
        {
            int intervalDevisor = 10;
            return (int?)Math.Ceiling((yMax - yMin) / intervalDevisor);
        
        }

        public static void NFCalculateNiceRange(double yMin, double yMax, int targetDivisions,
                                      out double newMin, out double newMax, out double interval)
        {
            double subMax = Math.Ceiling(yMax);
            interval = (int)Math.Ceiling((decimal)(subMax / 6));
            newMin = 0.0;

            newMax = 10 * interval;
            //double range = yMax - yMin;
            //double totalRange = newMax;
            while(newMax / yMax > 3)
            {
                if (interval <= 0.5)
                    break;
                interval = interval / 2.0;
                newMax = 10 * interval;
                
            }

        }

    }

    public enum LegendTextType
    { 
        Temp,
        Elec
    }

    public class InequalitySolver
    {
        // 定义不等式函数 F(x) = [yMin - x] - (2*(yMax - yMin) + 10*x)
        // 目标是找到 F(x) >= 0 的最大 x
        private static double CheckInequality(double yMin, double yMax, double x)
        {
            // 左侧: [yMin - x]
            double leftSide = Math.Floor(yMin - x);

            // 右侧: 2 * (yMax - yMin) + 10 * x
            double rightSide = 2 * (yMax - yMin) + 10 * x;

            // 检查 [yMin - x] >= rightSide 是否成立
            // 返回差值，如果差值 >= 0，则不等式成立
            return leftSide - rightSide;
        }

        /// <summary>
        /// 求解不等式 [yMin - x] >= 2(yMax - yMin) + 10x 的最大 x 值。
        /// </summary>
        /// <param name="yMin">yMin 值。</param>
        /// <param name="yMax">yMax 值。</param>
        /// <param name="precision">所需精度。</param>
        /// <returns>满足不等式的最大 x 值。</returns>
        public static double SolveForMaxX(double yMin, double yMax, double precision = 0.000001)
        {
            // 1. 计算近似解作为搜索的起点 (忽略取整)
            // x_approx = (3*yMin - 2*yMax) / 11
            double xApprox = (3 * yMin - 2 * yMax) / 11.0;

            // 2. 使用二分查找或简单线性搜索来精确定位边界

            // 检查 xApprox 是否满足条件 (取整的影响)
            if (CheckInequality(yMin, yMax, xApprox) >= 0)
            {
                // 如果满足，我们从 xApprox 开始递增（向正无穷）找到不满足的点
                // 然后再回到上一个点。
                double currentX = xApprox;
                double step = precision;

                // 搜索不满足条件的边界 (向右搜索)
                while (CheckInequality(yMin, yMax, currentX + step) >= 0)
                {
                    currentX += step;
                    // 增加 step，加快搜索速度，但要小心跳过边界
                    step *= 2;
                }

                // 找到大致边界后，进行精确二分查找
                double xLow = currentX;
                double xHigh = currentX + step;

                // 确保 xLow 满足条件
                if (CheckInequality(yMin, yMax, xLow) < 0)
                {
                    // 如果 xLow 不满足，从 xApprox 往回搜索 (不太可能发生，但为了健壮性)
                    xLow = xApprox;
                }

                // 在 [xLow, xHigh] 区间内进行二分查找
                while (xHigh - xLow > precision)
                {
                    double xMid = xLow + (xHigh - xLow) / 2.0;
                    if (CheckInequality(yMin, yMax, xMid) >= 0)
                    {
                        xLow = xMid; // 满足，继续向右搜索
                    }
                    else
                    {
                        xHigh = xMid; // 不满足，收缩到左侧
                    }
                }

                return xLow;
            }
            else
            {
                // 如果近似解不满足，说明 x 的最大值比 xApprox 小。
                // 从 xApprox 开始递减搜索（向负无穷）
                double currentX = xApprox;

                while (CheckInequality(yMin, yMax, currentX) < 0)
                {
                    currentX -= precision; // 线性递减搜索直到满足
                }
                return currentX;
            }
        }
    }
}

   

