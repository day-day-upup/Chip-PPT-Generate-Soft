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
            #region ��������
            public SeriesCollection LineSeriesCollection { get; set; } //SeriesCollection �� LiveCharts �ṩ���࣬���ڴ�Ŷ������ϵ��
            public Func<double, string> CustomFormatterX { get; set; } //��ʽ�� X ��ı�ǩ�������Զ�����ʾ�ĸ�ʽ
            public Func<double, string> CustomFormatterY { get; set; } //��ʽ�� Y ��ı�ǩ�������Զ�����ʾ�ĸ�ʽ

            private double axisXMax;
            public double AxisXMax //X��������ʾ��Χ
            {
                get { return axisXMax; }
                set { axisXMax = value; this.OnPropertyChanged("AxisXMax"); }
            }

            private double axisXMin;
            public double AxisXMin //X�����Сֵ
            {
                get { return axisXMin; }
                set { axisXMin = value; this.OnPropertyChanged("AxisXMin"); }
            }

            private double axisYMax;
            public double AxisYMax //Y��������ʾ��Χ
            {
                get { return axisYMax; }
                set
                {
                    axisYMax = value;
                    this.OnPropertyChanged("AxisYMax");
                }
            }
            private double axisYMin;
            public double AxisYMin //Y�����Сֵ
            {
                get { return axisYMin; }
                set
                {
                    axisYMin = value;
                    this.OnPropertyChanged("AxisYMin");
                }
            }

            private System.Timers.Timer timer = new System.Timers.Timer(); //����һ����ʱ��ʵ��
            private Random Randoms = new Random(); //�����������
            private int TabelShowCount = 10; //��ʾ��ͼ������ʾ��������   
            private List<ChartValues<double>> ValueLists { get; set; } //�洢 Y ������ݵ�
            private List<Axis> YAxes { get; set; } = new List<Axis>();

            private string CustomFormattersX(double val) //��ʽ�� X ��ı�ǩ
            {
                return string.Format("{0}", val); //���Գ�ʼ��Ϊʱ���
            }

            private string CustomFormattersY(double val) //��ʽ��  Y ��ı�ǩ
            {
                return string.Format("{0}", val);
            }
            #endregion

            public ChatViewModel()
            {
                AxisXMax = 10; //��ʼ��X������ֵΪ10
                AxisXMin = 0; //��ʼ��X�����СֵΪ0
                AxisYMax = 10; //��ʼ��Y������ֵΪ10
                AxisYMin = 0; //��ʼ��Y�����СֵΪ0

                ValueLists = new List<ChartValues<double>> // ��ʼ�������������ߵ�ֵ����
        {
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>(),
            new ChartValues<double>()
        };
                LineSeriesCollection = new SeriesCollection(); //����LineSeriesCollection��ʵ��

                CustomFormatterX = CustomFormattersX; //����X���Զ����ʽ������
                CustomFormatterY = CustomFormattersY; //����Y���Զ����ʽ������

                var colors = new[] //��ʼ��һ����ɫ���鹩������ɫʹ��
                {
                Brushes.Red,
                Brushes.Green,
                Brushes.Blue,
                Brushes.Orange,
                Brushes.Purple,
                Brushes.Teal
            };

                for (int i = 0; i < 6; i++) // �������� LineSeries ����ӵ�LineSeriesCollection������
                {
                    LineSeries lineseries = new LineSeries
                    {
                        Title = $"���� {i + 1}", // ������������
                        DataLabels = false, // ���ݱ�ǩ���ɼ�
                        Values = ValueLists[i], // Y ����ֵ�󶨵���Ӧ�� ValueList
                        StrokeThickness = 3, // ���������Ŀ��
                        PointGeometrySize = 8, // �������ݵ�Ĵ�С
                        LineSmoothness = 0.5, // �������ߵ����۶� (0: ֱ��, 1: ��ȫƽ��)
                        Stroke = colors[i % colors.Length], // ����ÿ�����ߵ���ɫ
                        Fill = Brushes.Transparent // ȥ����Ӱ
                    };
                    LineSeriesCollection.Add(lineseries); // ��ӵ� LineSeriesCollection ��
                }

                timer.Interval = 1000; //���ö�ʱ�����Ϊ1000���룬��1�봥��һ�ζ�ʱ�����ĵ��¼�
                timer.Enabled = false; //��ʱ����ʼδ�򿪣���Ҫ�ֶ���
                timer.Elapsed += �򿪶�ʱ����; //��ʱ���򿪶��ĵ��¼�
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit; //��ʱ���رն��ĵ��¼�
            AddData();
            AddData();
            }

            [RelayCommand]
            private void AddData()
            {
                for (int i = 0; i < ValueLists.Count; i++)  // Ϊÿ����������һ������� Y ֵ�����
                {
                    int yValue = Randoms.Next(2, 1000); // ���������
                    ValueLists[i].Add(yValue); // ���Ӧ�������������
                }

                int maxY = (int)ValueLists.Select(v => v.Max()).Max(); // ��ȡ�������ߵ����ֵ
                AxisYMax = maxY + 30; // �� Y ������ֵ����Ϊ������ֵ���� 30

                if (ValueLists[0].Count > TabelShowCount) // ������һ�����ߵ����ݵ�����
                {
                    AxisXMax = ValueLists[0].Count - 1; // X �����ֵ
                    AxisXMin = ValueLists[0].Count - TabelShowCount; // X ����Сֵ
                }
            }

            [RelayCommand]
            private void AddDataTimer()
            {
                if (timer.Enabled == false) //�ж϶�ʱ���Ƿ��Ǵ�״̬�����û�д򿪾ʹ򿪶�ʱ��������ݣ�
                {
                    timer.Start();
                }
                else
                {
                    timer.Stop(); //����Ѿ��򿪶�ʱ������ô��ε����ť���ǹرն�ʱ����ֹͣ�������
                    timer.Dispose();
                }
            }

            private void �򿪶�ʱ����(object sender, ElapsedEventArgs e) //��ʱ���򿪺��ĵ��¼�
            {
                for (int i = 0; i < ValueLists.Count; i++)  // Ϊÿ����������һ������� Y ֵ�����
                {
                    int yValue = Randoms.Next(2, 1000); // ���������
                    ValueLists[i].Add(yValue); // ���Ӧ�������������
                }

                int maxY = (int)ValueLists.Select(v => v.Max()).Max(); // ��ȡ�������ߵ����ֵ
                AxisYMax = maxY + 30; // �� Y ������ֵ����Ϊ������ֵ���� 30

                if (ValueLists[0].Count > TabelShowCount) // ������һ�����ߵ����ݵ�����
                {
                    AxisXMax = ValueLists[0].Count - 1; // X �����ֵ
                    AxisXMin = ValueLists[0].Count - TabelShowCount; // X ����Сֵ
                }
            }

            private void OnProcessExit(object sender, EventArgs e) //�����˳�ʱ���ĵ��¼�
            {
                try
                {
                    timer.Stop(); //�رն�ʱ��
                    timer.Dispose(); //�ͷ���Դ
                }
                catch { }
            }
        }
    
}
