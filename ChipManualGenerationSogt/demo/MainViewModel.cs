using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using LiveCharts;
using LiveCharts.Wpf;
using System.Timers;
using System.Windows.Media;
using System;
using System.Collections.Generic;
using System.Linq;
namespace demo
{
    

  
        public partial class ChatViewModel : ObservableObject
        {
            #region 属性声明
            public SeriesCollection LineSeriesCollection { get; set; } //SeriesCollection 是 LiveCharts 提供的类，用于存放多个数据系列
            public Func<double, string> CustomFormatterX { get; set; } //格式化 X 轴的标签。可以自定义显示的格式
            public Func<double, string> CustomFormatterY { get; set; } //格式化 Y 轴的标签。可以自定义显示的格式

            private double axisXMax;
            public double AxisXMax //X轴的最大显示范围
            {
                get { return axisXMax; }
                set { axisXMax = value; this.OnPropertyChanged("AxisXMax"); }
            }

            private double axisXMin;
            public double AxisXMin //X轴的最小值
            {
                get { return axisXMin; }
                set { axisXMin = value; this.OnPropertyChanged("AxisXMin"); }
            }

            private double axisYMax;
            public double AxisYMax //Y轴的最大显示范围
            {
                get { return axisYMax; }
                set
                {
                    axisYMax = value;
                    this.OnPropertyChanged("AxisYMax");
                }
            }
            private double axisYMin;
            public double AxisYMin //Y轴的最小值
            {
                get { return axisYMin; }
                set
                {
                    axisYMin = value;
                    this.OnPropertyChanged("AxisYMin");
                }
            }

            private System.Timers.Timer timer = new System.Timers.Timer(); //声明一个定时器实例
            private Random Randoms = new Random(); //随机数生成器
            private int TabelShowCount = 10; //表示在图表中显示的最大点数   
            private List<ChartValues<double>> ValueLists { get; set; } //存储 Y 轴的数据点
            private List<Axis> YAxes { get; set; } = new List<Axis>();

            private string CustomFormattersX(double val) //格式化 X 轴的标签
            {
                return string.Format("{0}", val); //可以初始化为时间等
            }

            private string CustomFormattersY(double val) //格式化  Y 轴的标签
            {
                return string.Format("{0}", val);
            }
            #endregion

            public ChatViewModel()
            {
                AxisXMax = 10; //初始化X轴的最大值为10
                AxisXMin = 0; //初始化X轴的最小值为0
                AxisYMax = 10; //初始化Y轴的最大值为10
                AxisYMin = 0; //初始化Y轴的最小值为0

                ValueLists = new List<ChartValues<double>> // 初始化六个数据曲线的值集合
        {
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>()
        };
                LineSeriesCollection = new SeriesCollection(); //创造LineSeriesCollection的实例

                CustomFormatterX = CustomFormattersX; //设置X轴自定义格式化函数
                CustomFormatterY = CustomFormattersY; //设置Y轴自定义格式化函数

                var colors = new[] //初始化一个颜色数组供线条颜色使用
                {
                Brushes.Red,
                Brushes.Green,
                Brushes.Blue,
                Brushes.Orange,
                Brushes.Purple,
                Brushes.Teal
            };

                for (int i = 0; i < 6; i++) // 创建六个 LineSeries 并添加到LineSeriesCollection集合中
                {
                    LineSeries lineseries = new LineSeries
                    {
                        Title = $"曲线 {i + 1}", // 设置曲线名称
                        DataLabels = false, // 数据标签不可见
                        Values = ValueLists[i], // Y 轴数值绑定到相应的 ValueList
                        StrokeThickness = 3, // 设置线条的宽度
                        PointGeometrySize = 8, // 设置数据点的大小
                        LineSmoothness = 0.5, // 设置折线的弯折度 (0: 直线, 1: 完全平滑)
                        Stroke = colors[i % colors.Length], // 设置每条曲线的颜色
                        Fill = Brushes.Transparent // 去掉阴影
                    };
                    LineSeriesCollection.Add(lineseries); // 添加到 LineSeriesCollection 中
                }

                timer.Interval = 1000; //设置定时器间隔为1000毫秒，即1秒触发一次定时器订阅的事件
                timer.Enabled = false; //定时器初始未打开，需要手动打开
                timer.Elapsed += 打开定时器了; //定时器打开订阅的事件
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit; //定时器关闭订阅的事件
            AddData();
            AddData();
            }

            [RelayCommand]
            private void AddData()
            {
                for (int i = 0; i < ValueLists.Count; i++)  // 为每条曲线生成一个随机的 Y 值并添加
                {
                    int yValue = Randoms.Next(2, 1000); // 生成随机数
                    ValueLists[i].Add(yValue); // 向对应的曲线添加数据
                }

                int maxY = (int)ValueLists.Select(v => v.Max()).Max(); // 获取所有曲线的最大值
                AxisYMax = maxY + 30; // 将 Y 轴的最大值设置为这个最大值加上 30

                if (ValueLists[0].Count > TabelShowCount) // 仅检查第一条曲线的数据点数量
                {
                    AxisXMax = ValueLists[0].Count - 1; // X 轴最大值
                    AxisXMin = ValueLists[0].Count - TabelShowCount; // X 轴最小值
                }
            }

            [RelayCommand]
            private void AddDataTimer()
            {
                if (timer.Enabled == false) //判断定时器是否是打开状态，如果没有打开就打开定时器添加数据；
                {
                    timer.Start();
                }
                else
                {
                    timer.Stop(); //如果已经打开定时器，那么这次点击按钮就是关闭定时器，停止添加数据
                    timer.Dispose();
                }
            }

            private void 打开定时器了(object sender, ElapsedEventArgs e) //定时器打开后订阅的事件
            {
                for (int i = 0; i < ValueLists.Count; i++)  // 为每条曲线生成一个随机的 Y 值并添加
                {
                    int yValue = Randoms.Next(2, 1000); // 生成随机数
                    ValueLists[i].Add(yValue); // 向对应的曲线添加数据
                }

                int maxY = (int)ValueLists.Select(v => v.Max()).Max(); // 获取所有曲线的最大值
                AxisYMax = maxY + 30; // 将 Y 轴的最大值设置为这个最大值加上 30

                if (ValueLists[0].Count > TabelShowCount) // 仅检查第一条曲线的数据点数量
                {
                    AxisXMax = ValueLists[0].Count - 1; // X 轴最大值
                    AxisXMin = ValueLists[0].Count - TabelShowCount; // X 轴最小值
                }
            }

            private void OnProcessExit(object sender, EventArgs e) //进程退出时订阅的事件
            {
                try
                {
                    timer.Stop(); //关闭定时器
                    timer.Dispose(); //释放资源
                }
                catch { }
            }
        }
    
}
