
using ScottPlot;
using ScottPlot.Interactivity;
using ScottPlot.WPF;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ChipManualGenerationSogt.models;
using static ChipManualGenerationSogt.Curves;
namespace ChipManualGenerationSogt
{
    /// <summary>
    /// Curves.xaml �Ľ����߼�
    /// </summary>
    public partial class Curves : UserControl
    {
        CurvesModel vm;
        S2PParser analyze;
        List<string> _curvesImgPath = new List<string>();
        
        public class XYPoint
        {
            public long Size { set; get; }
            public double[] XArrys { get; set; }
            public double[] YArrys { get; set; }
        }
        public class PlotParameters
        {
            public Collection<XYPoint> Points { set; get; }

            public string XLabel { set; get; }
            public string YLabel { set; get; }
            public string Title { set; get; }
            public Collection<string> LengenTexts { set; get; }

        }

        public Curves()
        {
            InitializeComponent();
            vm = new CurvesModel();
            this.DataContext = vm;
            
            //var analyzer = new S2PParser();
            //bool success = analyzer.Parse(@"F:\PROJECT\ChipManualGeneration\KP414-11TAPE_RP2DR1-25_C10R6-8-VD=4.6V&ID=79mA_2025-06-12 19.33.41_25.0deg_SPara.s2p");

            //var points = new Collection<XYPoint>();
            //var point1 = SPGenerateXYPointData(analyzer.S11, 0);
            //var point2 = SPGenerateXYPointData(analyzer.S22, 0);
            //var point3 = SPGenerateXYPointData(analyzer.S12, 0);
            //var point4 = SPGenerateXYPointData(analyzer.S21, 0);
            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //points.Add(point4);

            //var legends = new Collection<string>();
            //legends.Add("S11");
            //legends.Add("S22");
            //legends.Add("S12"); 
            //legends.Add("S21");

            //var plotParameters = new PlotParameters();
            //plotParameters.Points = points;
            //plotParameters.XLabel = "Frequency(GHz)";
            //plotParameters.YLabel = "Magnitude(dB)";
            //plotParameters.Title = "S Parameters";
            //plotParameters.LengenTexts = legends;
            //AddPlot(plotParameters);

            //var analyzer = new S2PParser();
            //analyzer.Parse(@"F:\PROJECT\ChipManualGeneration\ԭʼ����\MML004X_V2-����\MML004X_V2-25\-VD=4V&ID=67mA\L004X��L024X��L026X_MML004X_V2-25_-VD=4V&ID=67mA_2025-09-01 15.00.12_25.0deg_SPara.s2p");
            //var point = SPGenerateXYPointData(analyzer.S11, 0);
            //string yLable = "INPUT RETURN LOSS(dB)";
            //string xLable = "FREQUENCY(GHz)";
            //string title = "";
            //string legend = "VD=4.0V,IDQ=67mA";
            //AddPlot(GeneratePlotParameters(point, xLable, yLable, title, legend));

            //point = SPGenerateXYPointData(analyzer.S12, 0);
            //yLable = "ISOLATION(dB)";
            //AddPlot(GeneratePlotParameters(point, xLable, yLable, title, legend));

            //point = SPGenerateXYPointData(analyzer.S21, 0);
            //yLable = "Gain(dB)";
            //AddPlot(GeneratePlotParameters(point, xLable, yLable, title, legend));

            //point = SPGenerateXYPointData(analyzer.S22, 0);
            //yLable = "OUTPUT RETURN LOSS(dB)";
            //AddPlot(GeneratePlotParameters(point, xLable, yLable, title, legend));


            //var analyzer2 = new S2PParser();
            //analyzer2.Parse(@"F:\PROJECT\ChipManualGeneration\ԭʼ����\MML004X_V2-����\MML004X_V2-25\-VD=3V&ID=43mA\L004X��L024X��L026X_MML004X_V2-25_-VD=3V&ID=43mA_2025-09-01 14.49.13_25.0deg_SPara.s2p");

            //var analyzer3 = new S2PParser();
            //analyzer3.Parse(@"F:\PROJECT\ChipManualGeneration\ԭʼ����\MML004X_V2-����\MML004X_V2-25\-VD=5V&ID=90mA\L004X��L024X��L026X_MML004X_V2-25_-VD=5V&ID=90mA_2025-09-01 15.05.38_25.0deg_SPara.s2p");


            //var points = new Collection<XYPoint>();
            //var point1 = SPGenerateXYPointData(analyzer.S11, 0);
            //var point2 = SPGenerateXYPointData(analyzer2.S11, 0);
            //var point3 = SPGenerateXYPointData(analyzer3.S11, 0);

            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "INPUT RETURN LOSS(dB)";
            //var legends = new Collection<string>();
            //legends.Add("-VD=4V&ID=67mA");
            //legends.Add("-VD=3V&ID=43mA");
            //legends.Add("-VD=5V&ID=90mA");
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //points.Clear();
            //point1 = SPGenerateXYPointData(analyzer.S12, 0);
            //point2 = SPGenerateXYPointData(analyzer2.S12, 0);
            //point3 = SPGenerateXYPointData(analyzer3.S12, 0);
            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "ISOLATION(dB)";
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //points.Clear();
            //point1 = SPGenerateXYPointData(analyzer.S21, 0);
            //point2 = SPGenerateXYPointData(analyzer2.S21, 0);
            //point3 = SPGenerateXYPointData(analyzer3.S21, 0);
            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "GAIN(dB)";
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //points.Clear();
            //point1 = SPGenerateXYPointData(analyzer.S22, 0);
            //point2 = SPGenerateXYPointData(analyzer2.S22, 0);
            //point3 = SPGenerateXYPointData(analyzer3.S22, 0);
            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "OUTPUT RETURN LOSS(dB)";
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //int index = 0;
            //foreach(var item in vm.Plots)
            //{
            //    //item.Refresh();
            //    //Console.WriteLine(vm.Plots.Count);
            //    //string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            //    //string fileName = $"myplot_{index}.png";
            //    //item.Plot.SavePng(fileName, 600, 500);
            //    //item.Refresh();

            //    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            //    string folder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pic");
            //    Directory.CreateDirectory(folder);
            //    string fileName = System.IO.Path.Combine(folder, $"{index}.png");

            //    Application.Current.Dispatcher.Invoke(() =>
            //    {
            //        item.Plot.SavePng(fileName, 600, 500);
            //    });
            //    index++;
            //}
            //ScottPlot.Colors.Green;
            //curve.LinePattern = LinePattern.DenselyDashed;
            //AddPlot(plotParameters);
            //AddPlot();
        }
        public void AddPlot(PlotParameters plotParameters)
        {
            var plot = new WpfPlot();
            plot.Plot.Axes.Bottom.MinorTickStyle.Length = 0;//�����ӿ̶�
            plot.Plot.Axes.Left.MinorTickStyle.Length = 0;
            //plot.Plot.Axes.Bottom.TickGenerator = new ScottPlot.TickGenerators.NumericFixedInterval(1);// ÿ���̶ȵļ��
            //plot.Plot.Axes.Left.TickGenerator = new ScottPlot.TickGenerators.NumericFixedInterval(5);
            plot.Plot.Grid.MajorLineColor = ScottPlot.Colors.Black;
            //plot.MinWidth = plot.Height;

            // ����ͼ��
            //double[] dataX = { 1, 2, 3, 4, 5 };
            //double[] dataY = { 1, 4, 9, 16, 25 };
            if (plotParameters.Points.Count == 0)
                return;
            if (plotParameters.Points.Count != plotParameters.LengenTexts.Count)
                return;
            plot.Plot.Title(plotParameters.Title);
            plot.Plot.XLabel(plotParameters.XLabel);
            plot.Plot.YLabel(plotParameters.YLabel);
            for (int i = 0; i < plotParameters.Points.Count; i++)
            {
                var point = plotParameters.Points[i];
                var sig = plot.Plot.Add.SignalXY(point.XArrys, point.YArrys);
                sig.LegendText = plotParameters.LengenTexts[i];
                if (i == 0)
                {
                    sig.LinePattern = LinePattern.DenselyDashed;
                    sig.Color = ScottPlot.Colors.DeepSkyBlue;
                }
                //sig.LinePattern = LinePattern.DenselyDashed;
                if (i == 1)
                {
                    //sig.LinePattern = LinePattern.Solid;
                    sig.LinePattern = LinePattern.Solid;
                    sig.Color = ScottPlot.Colors.LimeGreen;
                    //var custom = new LinePattern(new float[] { 2, 2, 5, 10 }, 0, "Custom");
                    //sig.LinePattern = custom;

                }

                if (i == 2)
                {
                    sig.LinePattern = LinePattern.Dotted;
                    sig.Color = ScottPlot.Colors.Red;
                }
                if (i == 3)
                {
                    sig.LinePattern = LinePattern.Dashed;
                }

                if (plotParameters.Points.Count == 1)
                {
                    sig.LinePattern = LinePattern.Solid;
                    sig.Color = ScottPlot.Colors.Red;
                }

            }
            //plot.Plot.Legend();
            //plot.Refresh();
            //vm.Plots.Add(plot);


            //var points = plotParameters.Points;
            //var sig =plot.Plot.Add.SignalXY(point.XArrys, point.YArrys);

            //WpfPlot1.Plot.Legend.Alignment = ScottPlot.Alignment.UpperLeft;
            // sig1.Color = ScottPlot.Color.FromHex("#ff0000"); // �Ⱥ�ɫ

            //�һ�Ĭ�ϲ˵�ȫ�����
            plot.Menu.Clear();

            //�����Լ�����Ĳ˵� �һ�plotʱ
            //plot.Menu.Add("Save Png", plot1 =>
            //{
            //    plot1.SavePng("myplot.png", 600, 500); // ���ֱ����exeĿ¼������ͼƬ
            //    plot.Refresh(); // ˢ����ʾ
            //});
            //plot.Menu.Add("Change Data Source", plot1 =>
            //{
            //    //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
            //    //plot.Refresh(); // ˢ����ʾ
            //});
            plot.Menu.Add("Legend To UpperLeft", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.UpperLeft;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To UpperCenter", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.UpperCenter;
                plot.Refresh(); // ˢ����ʾ
            });

            plot.Menu.Add("Legend To UpperRight", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.UpperRight;
                plot.Refresh(); // ˢ����ʾ
            });

            plot.Menu.Add("Legend To LowerLeft", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.LowerLeft;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To LowerCenter", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.LowerCenter;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To LowerRight", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.LowerRight;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To MiddleLeft", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.MiddleLeft;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To MiddleCenter", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.MiddleCenter;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To MiddleRight", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.MiddleRight;
                plot.Refresh(); // ˢ����ʾ
            });

            Application.Current.Dispatcher.Invoke(() =>
            {
                plot.Refresh();  // �����߳�ִ�� UI ����
            });


            // ���ߴ�仯ʱ������ MinWidth = ActualHeight����֮��
            plot.SizeChanged += (s, e) =>
            {
                // ����1������С������ٵ��ڵ�ǰ�߶ȣ���ֹ̫խ��
                plot.MinWidth = Math.Max(400, plot.ActualHeight);

                // ����2������ǿ����С�����ȣ�ȡ�ϴ�ֵ��
                // double maxSize = Math.Max(plot.ActualWidth, plot.ActualHeight);
                // plot.MinWidth = plot.MinHeight = maxSize;
            };
            vm.Plots.Add(plot);
        }

        /// <summary>
        /// ����µ�ͼ
        /// </summary>
        /// <param name="plotModel">ͼ��ģ��
        public void AddPlot(PlotModel plotModel)
        {
            var plot = new WpfPlot();
            plot.Plot.Axes.Bottom.MinorTickStyle.Length = 0;//�����ӿ̶�
            plot.Plot.Axes.Left.MinorTickStyle.Length = 0;
            plot.Plot.Axes.Bottom.MajorTickStyle.Length = 0;
            plot.Plot.Axes.Left.MajorTickStyle.Length = 0;
            plot.Plot.Axes.Left.MajorTickStyle.Color = ScottPlot.Colors.Black;
            plot.Plot.Axes.Left.TickLabelStyle.FontSize = 25;
            plot.Plot.Axes.Left.TickLabelStyle.Bold = true;
            plot.Plot.Axes.Bottom.TickLabelStyle.FontSize = 25;
            plot.Plot.Axes.Bottom.TickLabelStyle.Bold = true;
            plot.Plot.Axes.Left.Label.FontSize = 28;
            plot.Plot.XLabel(plotModel.XLabel);
            plot.Plot.YLabel(plotModel.YLabel);
            if (plotModel.xAxisInterval != null && plotModel.yAxisInterval != null)
            {
                // ����x��y������̶ȼ��
                plot.Plot.Axes.Bottom.TickGenerator = new ScottPlot.TickGenerators.NumericFixedInterval((int)plotModel.xAxisInterval);// ÿ���̶ȵļ��
                plot.Plot.Axes.Left.TickGenerator = new ScottPlot.TickGenerators.NumericFixedInterval((int)plotModel.yAxisInterval);
            }
            plot.Plot.Grid.MajorLineColor = ScottPlot.Colors.Black;
            //plot.UserInputProcessor.IsEnabled = false;
            plot.UserInputProcessor.LeftClickDragPan(false, false, false);
            //RightClickDragZoom(false, bool horizontal, bool vertical)
            //if ( plotModel.yMin != null && plotModel.yMax != null)
            //    plot.Plot.Axes.SetLimitsY((double)plotModel.yMin, (double)plotModel.yMax); // �ֱ�����x��y�����С���ֵ
            //else if (plotModel.xMin != null && plotModel.xMax != null)
            //{
            //    plot.Plot.Axes.SetLimitsX((double)plotModel.xMin, (double)plotModel.xMax);
            //}
            plot.Plot.Axes.SetLimits(plotModel.xMin, plotModel.xMax, plotModel.yMin, plotModel.yMax);
            plot.Plot.Legend.Alignment = plotModel.Alignment;
            var styles = new (ScottPlot.Color Color, ScottPlot.LinePattern Pattern)[]
                         {
                             (ScottPlot.Colors.Green, ScottPlot.LinePattern.Solid),
                             (ScottPlot.Colors.Xkcd.BrightRed, ScottPlot.LinePattern.Dotted),


                            (ScottPlot.Colors.Xkcd.BrightBlue, ScottPlot.LinePattern.Dashed), // i == 0
                                     // i == 1
                                          // i == 2
                            
                                                                                                // ���������������Ӹ���Ԥ�������ʽ
                         };

            // ��鼯���Ƿ�ֻ��һ��Ԫ��
            bool isSingleCurve = plotModel.Cureves.Count == 1;

            for (int i = 0; i < plotModel.Cureves.Count; i++)
            {
                var curve = plotModel.Cureves.ElementAt(i);

                // 1. ��ȡ���ݲ����� SignalXY ��ͼ����
                var sig = plot.Plot.Add.SignalXY(curve.XData, curve.YData);
                sig.LegendText = curve.Legend;
                sig.LineWidth = 5;
                //sig.LineStyle.
                // 2. ������������Ӧ����ʽ�߼�
                if (isSingleCurve)
                {
                    // ���ֻ��һ�����ߣ�Ӧ���ض���Ĭ����ʽ
                    sig.LinePattern = ScottPlot.LinePattern.Solid;
                    sig.Color = ScottPlot.Colors.Red;
                }
                else
                {
                    // ����ж������ߣ���Ԥ������ʽ�л�ȡ��ʹ��ȡģ������ѭ����ʽ��
                    int styleIndex = i % styles.Length;
                    sig.Color = styles[styleIndex].Color;
                    sig.LinePattern = styles[styleIndex].Pattern;
                }
            }

            //�һ�Ĭ�ϲ˵�ȫ�����
            plot.Menu.Clear();

            plot.Menu.Add("Legend To UpperLeft", plot1 =>
            {
                //plot1.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.UpperLeft;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To UpperCenter", plot1 =>
            {

                plot.Plot.Legend.Alignment = ScottPlot.Alignment.UpperCenter;
                plot.Refresh(); // ˢ����ʾ
            });

            plot.Menu.Add("Legend To UpperRight", plot1 =>
            {

                plot.Plot.Legend.Alignment = ScottPlot.Alignment.UpperRight;
                plot.Refresh(); // ˢ����ʾ
            });

            plot.Menu.Add("Legend To LowerLeft", plot1 =>
            {

                plot.Plot.Legend.Alignment = ScottPlot.Alignment.LowerLeft;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To LowerCenter", plot1 =>
            {

                plot.Plot.Legend.Alignment = ScottPlot.Alignment.LowerCenter;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To LowerRight", plot1 =>
            {

                plot.Plot.Legend.Alignment = ScottPlot.Alignment.LowerRight;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To MiddleLeft", plot1 =>
            {

                plot.Plot.Legend.Alignment = ScottPlot.Alignment.MiddleLeft;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To MiddleCenter", plot1 =>
            {

                plot.Plot.Legend.Alignment = ScottPlot.Alignment.MiddleCenter;
                plot.Refresh(); // ˢ����ʾ
            });
            plot.Menu.Add("Legend To MiddleRight", plot1 =>
            {
                plot.Plot.Legend.Alignment = ScottPlot.Alignment.MiddleRight;
                plot.Refresh(); // ˢ����ʾ
            });

            Application.Current.Dispatcher.Invoke(() =>
            {
                plot.Refresh();  // �����߳�ִ�� UI ����
            });


            // ���ߴ�仯ʱ������ MinWidth = ActualHeight����֮��
            plot.SizeChanged += (s, e) =>
            {
                // ����1������С������ٵ��ڵ�ǰ�߶ȣ���ֹ̫խ��
                plot.MinWidth = Math.Max(400, plot.ActualHeight);

                // ����2������ǿ����С�����ȣ�ȡ�ϴ�ֵ��
                // double maxSize = Math.Max(plot.ActualWidth, plot.ActualHeight);
                // plot.MinWidth = plot.MinHeight = maxSize;
            };
            vm.Plots.Add(plot);
        }
        public void Clear()
        {
            vm.Plots.Clear();
        
        }

        public S2PParser PaserS2p(string filePath)
        {
            var analyzer = new S2PParser();
            bool success = analyzer.Parse(filePath);
            if (success)
            {
                Console.WriteLine($"��ȡ {analyzer.S11.Count} �� S11 ���ݵ�");
                foreach (var s in analyzer.S11.Take(20))
                {
                    Console.WriteLine($"Freq: {s.FreqGHz:F3} GHz, S11: {s.DbValue:F2} dB");
                }
            }
            else
            {
                Console.WriteLine("��ȡʧ�ܣ�");
            }

            return analyzer;
        }

        public XYPoint SPGenerateXYPointData(List<S_Parameters> sParameters, int type)
        {
            // type ȷ����s�������ݽṹ�е�db ����phase
            var xyPouint = new XYPoint();
            List<double> xData = new List<double>();
            List<double> yData = new List<double>();
            foreach (var s in sParameters)
            {
                xData.Add(s.FreqGHz);
            }
            if (type == 0)// ʹ��db����
            {

                foreach (var s in sParameters)
                {
                    yData.Add(s.DbValue);
                }
            }
            else // ʹ��phase����
            {
                foreach (var s in sParameters)
                {
                    yData.Add(s.PhaseValue);
                }
            }
            xyPouint.XArrys = xData.ToArray();
            xyPouint.YArrys = yData.ToArray();
            xyPouint.Size = sParameters.Count;

            return xyPouint;
        }

        public PlotParameters GeneratePlotParameters(XYPoint xyPoint, string xLabel, string yLabel, string title, string legends)
        {
            var plotParameters = new PlotParameters();
            var points = new Collection<XYPoint>();
            points.Add(xyPoint);
            var lengends = new Collection<string>();
            lengends.Add(legends);
            plotParameters.Points = points;
            plotParameters.XLabel = xLabel;
            plotParameters.YLabel = yLabel;
            plotParameters.Title = title;
            plotParameters.LengenTexts = lengends;

            return plotParameters;
        }

        public XYPoint SPGenerateXYPointData(List<Pair_Parameters> sParameters)
        {
            // type ȷ����s�������ݽṹ�е�db ����phase
            var xyPouint = new XYPoint();
            List<double> xData = new List<double>();
            List<double> yData = new List<double>();
            foreach (var s in sParameters)
            {
                xData.Add(s.FreqGHz);
            }


            foreach (var s in sParameters)
            {
                yData.Add(s.Value);
            }


            xyPouint.XArrys = xData.ToArray();
            xyPouint.YArrys = yData.ToArray();
            xyPouint.Size = sParameters.Count;

            return xyPouint;
        }
        public void DeleteAllPlot()
        {
            vm.Plots.Clear();

        }
        public PlotParameters GeneratePlotParameters(Collection<XYPoint> points, string xLabel, string yLabel, string title, Collection<string> legends)
        {
            var plotParameters = new PlotParameters();
            plotParameters.Points = points;
            plotParameters.XLabel = xLabel;
            plotParameters.YLabel = yLabel;
            plotParameters.Title = title;
            plotParameters.LengenTexts = legends;
            return plotParameters;
        }
        private void PlotChangeData(WpfPlot plot, XYPoint point)
        {
            plot.Plot.Clear();
            plot.Plot.Add.SignalXY(point.XArrys, point.YArrys);
        }

        public void SaveAllPlot(string filePath)
        {
            _curvesImgPath.Clear();
            int index = 0;
            foreach (var item in vm.Plots)
            {
                //item.Refresh();
                //Console.WriteLine(vm.Plots.Count);
                //string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                //string fileName = $"myplot_{index}.png";
                //item.Plot.SavePng(fileName, 600, 500);
                //item.Refresh();

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                string folder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "pic");
                Directory.CreateDirectory(folder);
                string fileName = System.IO.Path.Combine(folder, $"{index}.png");
                _curvesImgPath.Add(fileName);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    item.Plot.SavePng(fileName, 600, 500);
                });
                index++;
            }
        }

        public List<string> GetAllCurvesImagesFilePath()
        {
            
            return _curvesImgPath;
        }

        public ObservableCollection<WpfPlot> GetPlots()
        { 
        
            return vm.Plots;
        }
    }
    internal class CurvesModel : ObeservableObject
    {
        private WpfPlot plot;
        public WpfPlot Plot
        {
            get { return plot; }
            set { this.plot = value; RaisePropertyChanged(nameof(Plot)); }
        }
        public ObservableCollection<WpfPlot> Plots { set; get; } = new ObservableCollection<WpfPlot>();
    }

}
