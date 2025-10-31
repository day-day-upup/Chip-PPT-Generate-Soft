
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using OpenTK;
using ScottPlot;
using ScottPlot.Interactivity.UserActionResponses;
using ScottPlot.WPF;


using static OpenTK.Graphics.OpenGL.GL;
namespace demo
{
    /// <summary>
    /// MainWindow.xaml �Ľ����߼�
    /// </summary>
    public partial class MainWindow : Window
    {
        WorkViewModel vm;
        private int _currentGroupCount=1;
       
       
        public MainWindow()
        {
            InitializeComponent();
            vm = new WorkViewModel();
            this.DataContext = vm;
            //vm.PopupVisible = false;
            //vm.Records.Add(new NewTreeViewItem
            //{
            //    Content = "Amplifier",
            //    Visible = true,
            //   // Children.Add(new NewTreeViewItem { Content = "�����ձ�" }),
            //    Children =
            //        {
            //            new NewTreeViewItem{ Content = "MML806" },
            //            new NewTreeViewItem { Content = "MML807" }
            //        }
            //});
            //vm.Records.Add(new NewTreeViewItem
            //{
            //    Content = "Filter",
            //    Visible = true,
            //    Children =
            //        {
            //            new NewTreeViewItem{ Content = "MML806" },
            //            new NewTreeViewItem { Content = "MML807" }
            //        }
            //});
            //vm.ViewModel = new ChatViewModel();
            // var sig1 = WpfPlot1.Plot.Add.Signal(Generate.Sin(51));
            // sig1.Color = ScottPlot.Color.FromHex("#ff0000"); // �Ⱥ�ɫ
            // sig1.LegendText = "Sin";
            // WpfPlot1.Plot.YLabel("Voltage (V)");
            // WpfPlot1.Plot.XLabel("Time (s)");

            //var sig2 = WpfPlot1.Plot.Add.Signal(Generate.Cos(51));
            // sig2.LegendText = "Cos";
            // sig2.Color = ScottPlot.Color.FromHex("#00ff00"); // �Ⱥ�ɫ

            // WpfPlot1.Plot.ShowLegend();
            // WpfPlot1.Plot.Legend.Alignment = ScottPlot.Alignment.UpperLeft;
            // WpfPlot1.Menu.Clear();
            // //WpfPlot1.Menu.Add("Reset Zoom");
            // WpfPlot1.Menu.Add("Save Png", plot =>
            // {
            //     plot.SavePng("myplot.png", 600, 400); // ���ֱ����exeĿ¼������ͼƬ
            //     WpfPlot1.Refresh(); // ˢ����ʾ
            // });
            //var menu = WpfPlot1.ContextMenu;

            //// ɾ��ĳЩ��
            //menu.Items.RemoveAt(0); // �Ƴ���һ��
            //                        // ���߸��� Header ��
            //foreach (var item in menu.Items.OfType<MenuItem>().ToList())
            //{
            //    if (item.Header?.ToString() == "Help")
            //        menu.Items.Remove(item);
            //}

            //// ����Զ�����
            //menu.Items.Add(new MenuItem()
            //{
            //    Header = "�Զ��幦��",
            //    Command = new RoutedCommand(), // ������ Click �¼�
            //});
            //AddPlot();
            //AddPlot();
            //InitializeParameterRows(1);
            //RebuildDataGridColumns();
            //InitializeFeatureParameterRows();
            //FeatureRebuildDataGridColumns();
            //double[] sin = Generate.Sin(51);

            // add a signal plot to the plot
            //WpfPlot1.Plot.Add.Signal(sin);

            //WpfPlot1.Refre sh();

            var wpf = new WpfPlot();
            double[] dataX = { 1, 2, 3, 4, 5 };
            double[] dataY = { 1, 4, 9, 16, 25 };
            var sig =WpfPlot1.Plot.Add.Scatter(dataX, dataY);
            sig.LegendText = "132465";
            WpfPlot1.Plot.Axes.Bottom.MinorTickStyle.Length = 0;//�����ӿ̶�
            WpfPlot1.Plot.Axes.Left.MinorTickStyle.Length = 0;
            WpfPlot1.Plot.Axes.Bottom.TickGenerator = new ScottPlot.TickGenerators.NumericFixedInterval(1);// ÿ���̶ȵļ��
            WpfPlot1.Plot.Axes.Left.TickGenerator = new ScottPlot.TickGenerators.NumericFixedInterval(5);
            WpfPlot1.Plot.Grid.MajorLineColor = ScottPlot.Colors.Black;

            //var panResponse = new ScottPlot.Interactivity.UserActionResponses.MouseDragPan(panButton);
            WpfPlot1.UserInputProcessor.IsEnabled = false;

            WpfPlot1.Plot.Axes.SetLimits(0, 5, 0, 25);
            //WpfPlot1.Plot.Interactivity.UserActionResponses
            //ScottPlot.Interactivity.UserActionResponses
            //WpfPlot1.Plot.UserActionResponses.MouseDragPan.LockX = true;
            //wpfPlot.Interactivity.UserActionResponses.MouseDragPan.LockY = true;
            WpfPlot1.Refresh();

            
        }

        private void btn_Add_TreeItem_Click(object sender, RoutedEventArgs e)
        {
            var menu = sender as MenuItem;
            var item = menu?.DataContext as NewTreeViewItem;
            Console.WriteLine(123);
            if (item != null)
            {
                Console.WriteLine(item.Content); // ? ��ȫ����
                //item.Children.Add(new NewTreeViewItem { Content = "MML808" });
                //vm.Record = item;
                //vm.PopupVisible = true;
                // ��Ҳ���Ը�ֵ�� vm.Record�������Ҫ��
                //vm.Record = item;
            }
        }

        private void Popup_OK_Click(object sender, RoutedEventArgs e)
        {
            //vm.Record.Children.Add(new NewTreeViewItem { Content = vm.PopupText });
            //vm.PopupVisible = false;
        }

        private void Popup_Cancel_Click(object sender, RoutedEventArgs e)
        {
            //vm.PopupVisible = false;
        }

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var selectedItem = e.NewValue as NewTreeViewItem; // ?? ��󶨵���������

            if (selectedItem != null)
            {
                Console.WriteLine($"�����ˣ�{selectedItem.Content}");
                // �������������� ViewModel�������˵�����ʾ�����
            }
        }

        private void WpfPlot1_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
          
        }

        private void ResetZoom_Click(object sender, RoutedEventArgs e)
        {
            //WpfPlot1.Plot.AxisAuto();
            //WpfPlot1.Refresh();
        }

        private void SaveImage_Click(object sender, RoutedEventArgs e)
        {
            //WpfPlot1.Plot.SaveFig("myplot.png");
            //WpfPlot1.Plot.SavePng("myplot.png", 600, 400);
        }


       


        private void AddPlot()
        {
            var plot = new WpfPlot();
           
            //plot.MinWidth = plot.Height;

            // ����ͼ��
            //var data = ScottPlot.Generate.Sin(51 + vm.Plots.Count * 10);
            //plot.Plot.Add.Signal(data);
            //plot.Plot.Title($"Chart {vm.Plots.Count + 1}");
            
            //plot.Refresh();
            
            //// ���ߴ�仯ʱ������ MinWidth = ActualHeight����֮��
            //plot.SizeChanged += (s, e) =>
            //{
            //    // ����1������С������ٵ��ڵ�ǰ�߶ȣ���ֹ̫խ��
            //    plot.MinWidth = Math.Max(400, plot.ActualHeight);

            //    // ����2������ǿ����С�����ȣ�ȡ�ϴ�ֵ��
            //    // double maxSize = Math.Max(plot.ActualWidth, plot.ActualHeight);
            //    // plot.MinWidth = plot.MinHeight = maxSize;
            //};
            //vm.Plots.Add(plot);
        }

        private void RemovePlot()
        {
            //if (vm.Plots.Count > 0)
            //    vm.Plots.RemoveAt(vm.Plots.Count - 1);
        }

        // ��ť�¼�
        private void AddButton_Click(object sender, RoutedEventArgs e) => AddPlot();
        private void RemoveButton_Click(object sender, RoutedEventArgs e) => RemovePlot();



        private void InitializeParameterRows(int groupCount)
        {
            //vm.ParameterRows.Clear();

            // ʾ�����������������
            AddParameterRow("Frequency", "GHz", groupCount);
            AddParameterRow("Small Signal Gain", "dB", groupCount);
            AddParameterRow("Gain Flatness", "dB", groupCount);
            AddParameterRow("Noise Figure", "dB", groupCount);
            AddParameterRow("P1dB - Output 1dB Compression", "dBm", groupCount);
            AddParameterRow("Psat - Saturated  Output Power", "dBm", groupCount);
            AddParameterRow("OIP3 - Output Third Order Intercept", "dBm", groupCount);
            AddParameterRow("Input Return Loss", "dB", groupCount);
            AddParameterRow("Output Return Loss", "dB", groupCount);
        }

        private void InitializeFeatureParameterRows()
        {
            //vm.FeatureParameterRows.Clear();
            //var row = new FeatureParameterRow
            //{
            //    Name = "Frequency",
            //    Info = "45-90GHz"
            //};
            //vm.FeatureParameterRows.Add(row);
            FeatureAddParameterRow("Frequency", "45-90GHz");
            FeatureAddParameterRow("Small Signal Gain", "15dB Typical");
            FeatureAddParameterRow("Gain Flatness", "��2.5dB Typical");
            FeatureAddParameterRow("Noise Figure", "4.5dB Typical");
            FeatureAddParameterRow("P1dB", "12dBm Typical");
            FeatureAddParameterRow("Power Supply", "VD=+4V@119mA ,VG=-0.4V");
            FeatureAddParameterRow("Input/Output", "50��");
            FeatureAddParameterRow("Chip Size", "1.766 x 2.0 x 0.05mm ");
        }
        private void AddParameterRow(string name, string unit, int groupCount)
        {
            var row = new ParameterRow
            {
                Name = name,
                Unit = unit,
                Groups = new ObservableCollection<MinMaxTypeGroup>()
            };

            for (int i = 0; i < groupCount; i++)
            {
                row.Groups.Add(new MinMaxTypeGroup());
            }

            //vm.ParameterRows.Add(row);
        }

        private void FeatureAddParameterRow(string name, string info)
        {
            var row = new FeatureParameterRow
            {
                Name = name,
                Info = info
            };
            //vm.FeatureParameterRows.Add(row);
        }

        private void RebuildDataGridColumns()
        {
            //dataGrid.Columns.Clear();

            //// ��һ�У���������
            //dataGrid.Columns.Add(new DataGridTextColumn
            //{
            //    Header = "Parameters",
            //    Binding = new Binding("Name"),
            //    IsReadOnly = true
            //});

            //// ��̬��� Min/Type/Max ��
            //if (vm.ParameterRows.Count > 0)
            //{
            //    int groupCount = vm.ParameterRows[0].Groups.Count; // ��������������һ��

            //    for (int i = 0; i < groupCount; i++)
            //    {
            //        dataGrid.Columns.Add(new DataGridTextColumn
            //        {
            //            Header = $"Min {i + 1}",
            //            Binding = new Binding($"Groups[{i}].Min")
            //        });

            //        dataGrid.Columns.Add(new DataGridTextColumn
            //        {
            //            Header = $"Type {i + 1}",
            //            Binding = new Binding($"Groups[{i}].Type")
            //        });

            //        dataGrid.Columns.Add(new DataGridTextColumn
            //        {
            //            Header = $"Max {i + 1}",
            //            Binding = new Binding($"Groups[{i}].Max")
            //        });
            //    }
            //}

            //// ���һ�У���λ
            //dataGrid.Columns.Add(new DataGridTextColumn
            //{
            //    Header = "Units",
            //    Binding = new Binding("Unit"),
            //    IsReadOnly = true
            //});
        }

        private void FeatureRebuildDataGridColumns()
        {
            //featureDataGrid.Columns.Clear();

            //// ��һ�У���������
            //featureDataGrid.Columns.Add(new DataGridTextColumn
            //{
            //    Header = "Feature",
            //    Binding = new Binding("Name"),
            //    //IsReadOnly = true
            //});


            //// ���һ�У���λ
            //featureDataGrid.Columns.Add(new DataGridTextColumn
            //{
            //    Header = " ",
            //    Binding = new Binding("Info"),
            //    //IsReadOnly = true
            //});
        }

        private void AddGroup_Click(object sender, RoutedEventArgs e)
        {
            //_currentGroupCount++;
            //foreach (var row in vm.ParameterRows)
            //{
            //    row.Groups.Add(new MinMaxTypeGroup());
            //}
            //RebuildDataGridColumns();
        }

        private void RemoveGroup_Click(object sender, RoutedEventArgs e)
        {
            //if (_currentGroupCount > 1)
            //{
            //    _currentGroupCount--;
            //    foreach (var row in vm.ParameterRows)
            //    {
            //        row.Groups.RemoveAt(row.Groups.Count - 1);
            //    }
            //    RebuildDataGridColumns();
            //}
        }

        private void AddFeatureRow_Click(object sender, RoutedEventArgs e)
        {
            //vm.FeatureParameterRows.Add(new FeatureParameterRow());
        }
        private void InsertRowAbove_Click(object sender, RoutedEventArgs e)
        {
            InsertRowAtOffset(0);
        }
        // ��ѡ�����·�����
        private void InsertRowBelow_Click(object sender, RoutedEventArgs e)
        {
            InsertRowAtOffset(1); // ���뵽��ǰ���� + 0�����·���
        }
        // ͨ�ò��뷽��
        private void InsertRowAtOffset(int offset)
        {
            //if (featureDataGrid.SelectedItems.Count == 0)
            //    return;

            //// ��ȡ��һ��ѡ���֧�ֶ�ѡ����ֻȡ��һ����
            //var selectedItem = featureDataGrid.SelectedItems[0];

            //// �ҵ�����Դ�����е�����
            //int currentIndex = vm.FeatureParameterRows.IndexOf(selectedItem as FeatureParameterRow);
            //if (currentIndex == -1)
            //    return;

            //// �������λ��
            //int insertIndex = currentIndex + offset;

            //// �߽紦������С�� 0
            //if (insertIndex < 0)
            //    insertIndex = 0;

            //// ��������
            //vm.FeatureParameterRows.Insert(insertIndex, new FeatureParameterRow());

            //// ��ѡ���Զ�ѡ���²������
            //featureDataGrid.SelectedIndex = insertIndex;
            //featureDataGrid.ScrollIntoView(vm.FeatureParameterRows[insertIndex]);
        }
        private void DeleteFeatureRow_Click(object sender, RoutedEventArgs e)
        {
            // ��ȡ��ǰѡ�е�������Ƕ����
            //var selectedItems = new List<object>(featureDataGrid.SelectedItems.Cast<object>());

            //// �Ӻ���ǰɾ�������������仯���⣩
            //foreach (var item in selectedItems)
            //{
            //    if (item is FeatureParameterRow row)
            //    {
            //        vm.FeatureParameterRows.Remove(row);
            //    }
            //}
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ��ȡ�����¼��� ListBox
            var listBox = sender as ListBox;

            // ȷ����ѡ����
            if (listBox?.SelectedItem == null)
                return;

            // ��ȡѡ��������ݣ�Content��
            var selectedItem = listBox.SelectedItem as ListBoxItem;
            if (selectedItem == null)
                return;

            string content = selectedItem.Content?.ToString();

            // ���������жϵ�����ĸ�
            if (content == "Paramtets Table")  // ע�⣺��ƴд���� "Paramtets"������ "Parameters"
            {
                // �л��� Parameters ���
                ShowParametersTable();
            }
            else if (content == "Feature Table")
            {
                // �л��� Feature ���
                ShowFeatureTable();
            }
        }
        private void ShowParametersTable()
        {
            // ���磺��ʾ����������� feature ���
            //dataGrid.Visibility = Visibility.Visible;
            //featureDataGrid.Visibility = Visibility.Collapsed;
        }

        private void ShowFeatureTable()
        {
            // ��ʾ feature ������ز������
            //dataGrid.Visibility = Visibility.Collapsed;
            //featureDataGrid.Visibility = Visibility.Visible;
        }
    
    }

    public class TaskModel : INotifyPropertyChanged
    {
        public string ID { get; set; }
        public string TaskName { get; set; }
        public string Status { get; set; }
        public string Project { get; set; }
        public string Executor { get; set; }
        public string Deadline { get; set; }
        public string Estimated { get; set; }
        public string Consumed { get; set; }
        public string Remaining { get; set; }
        public string Creator { get; set; }
        public string Finisher { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    internal class WorkViewModel : ObeservableObject 
    {
        
        public WorkViewModel()
        {

        }

        /// <summary>
        /// Gets the plot model.
        /// </summary>
      



        public ObservableCollection<TaskModel> Tasks { get; set; } = new ObservableCollection<TaskModel>();
       


    }

    public class NewTreeViewItem : ObeservableObject
    {
        private string content;
        public string Content
        {
            get { return content; }
            set { this.content = value; RaisePropertyChanged(nameof(content)); }
        }


        private bool visible=false;
        public bool Visible
        {
            get { return visible; }
            set { this.visible = value; RaisePropertyChanged(nameof(visible)); }
        }
        public ObservableCollection<NewTreeViewItem> Children { get; set; } = new ObservableCollection<NewTreeViewItem>();
       
    }

    public class Entity : ObeservableObject 
    {
        private string email;
        public string Email
        {
            get { return email; }   
            set { this.email = value; RaisePropertyChanged(nameof(Email)); }
        }

        private string name;
        public string Name
        {
            get { return name; }
            set { this.name = value; RaisePropertyChanged(nameof(Name)); }
        }

        private string phone;
        public string Phone
        {
            get { return phone; }
            set { this.phone = value; RaisePropertyChanged(nameof(Phone)); }
        }

    }


    public class MinMaxTypeGroup : INotifyPropertyChanged
    {
        private double _min;
        private string _type;
        private double _max;

        public double Min
        {
            get => _min;
            set { _min = value; OnPropertyChanged(); }
        }

        public string Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(); }
        }

        public double Max
        {
            get => _max;
            set { _max = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var prop = GetType().GetProperty(propertyName);
            object value = prop?.GetValue(this) ?? "NULL";
            Debug.WriteLine($"Property {propertyName} changed to: {value}");
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ParameterRow : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string Unit { get; set; }
        public ObservableCollection<MinMaxTypeGroup> Groups { get; set; } = new ObservableCollection<MinMaxTypeGroup>();

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class FeatureParameterRow : INotifyPropertyChanged
    {
        public FeatureParameterRow()
        {
            // ȷ�� Name �� Info ��Ĭ��ֵ������ null ��ʾ���⣩
            Name = string.Empty;
            Info = string.Empty;
        }

        private string _info;

        public string Name
        {
            get;
            set; 
        }

        public string Info
        {
            get => _info;
            set { _info = value; OnPropertyChanged(); }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var prop = GetType().GetProperty(propertyName);
            object value = prop?.GetValue(this) ?? "NULL";
            Debug.WriteLine($"Property {propertyName} changed to: {value}");
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
