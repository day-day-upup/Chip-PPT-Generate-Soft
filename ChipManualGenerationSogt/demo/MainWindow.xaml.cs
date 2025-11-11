
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
    /// MainWindow.xaml 的交互逻辑
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
            //   // Children.Add(new NewTreeViewItem { Content = "人民日报" }),
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
            // sig1.Color = ScottPlot.Color.FromHex("#ff0000"); // 橙红色
            // sig1.LegendText = "Sin";
            // WpfPlot1.Plot.YLabel("Voltage (V)");
            // WpfPlot1.Plot.XLabel("Time (s)");

            //var sig2 = WpfPlot1.Plot.Add.Signal(Generate.Cos(51));
            // sig2.LegendText = "Cos";
            // sig2.Color = ScottPlot.Color.FromHex("#00ff00"); // 橙红色

            // WpfPlot1.Plot.ShowLegend();
            // WpfPlot1.Plot.Legend.Alignment = ScottPlot.Alignment.UpperLeft;
            // WpfPlot1.Menu.Clear();
            // //WpfPlot1.Menu.Add("Reset Zoom");
            // WpfPlot1.Menu.Add("Save Png", plot =>
            // {
            //     plot.SavePng("myplot.png", 600, 400); // 这个直接在exe目录下生成图片
            //     WpfPlot1.Refresh(); // 刷新显示
            // });
            //var menu = WpfPlot1.ContextMenu;

            //// 删除某些项
            //menu.Items.RemoveAt(0); // 移除第一个
            //                        // 或者根据 Header 找
            //foreach (var item in menu.Items.OfType<MenuItem>().ToList())
            //{
            //    if (item.Header?.ToString() == "Help")
            //        menu.Items.Remove(item);
            //}

            //// 添加自定义项
            //menu.Items.Add(new MenuItem()
            //{
            //    Header = "自定义功能",
            //    Command = new RoutedCommand(), // 或者用 Click 事件
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
            WpfPlot1.Plot.Axes.Bottom.MinorTickStyle.Length = 0;//禁用子刻度
            WpfPlot1.Plot.Axes.Left.MinorTickStyle.Length = 0;
            WpfPlot1.Plot.Axes.Bottom.TickGenerator = new ScottPlot.TickGenerators.NumericFixedInterval(1);// 每个刻度的间隔
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
                Console.WriteLine(item.Content); // ? 安全访问
                //item.Children.Add(new NewTreeViewItem { Content = "MML808" });
                //vm.Record = item;
                //vm.PopupVisible = true;
                // 你也可以赋值给 vm.Record（如果需要）
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
            var selectedItem = e.NewValue as NewTreeViewItem; // ?? 你绑定的数据类型

            if (selectedItem != null)
            {
                Console.WriteLine($"你点击了：{selectedItem.Content}");
                // 你可以在这里更新 ViewModel、弹出菜单、显示详情等
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

            // 配置图表
            //var data = ScottPlot.Generate.Sin(51 + vm.Plots.Count * 10);
            //plot.Plot.Add.Signal(data);
            //plot.Plot.Title($"Chart {vm.Plots.Count + 1}");
            
            //plot.Refresh();
            
            //// 当尺寸变化时，保持 MinWidth = ActualHeight（或反之）
            //plot.SizeChanged += (s, e) =>
            //{
            //    // 方法1：让最小宽度至少等于当前高度（防止太窄）
            //    plot.MinWidth = Math.Max(400, plot.ActualHeight);

            //    // 方法2：或者强制最小宽高相等（取较大值）
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

        // 按钮事件
        private void AddButton_Click(object sender, RoutedEventArgs e) => AddPlot();
        private void RemoveButton_Click(object sender, RoutedEventArgs e) => RemovePlot();



        private void InitializeParameterRows(int groupCount)
        {
            //vm.ParameterRows.Clear();

            // 示例：添加两个参数行
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
            FeatureAddParameterRow("Gain Flatness", "±2.5dB Typical");
            FeatureAddParameterRow("Noise Figure", "4.5dB Typical");
            FeatureAddParameterRow("P1dB", "12dBm Typical");
            FeatureAddParameterRow("Power Supply", "VD=+4V@119mA ,VG=-0.4V");
            FeatureAddParameterRow("Input/Output", "50Ω");
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

            //// 第一列：参数名称
            //dataGrid.Columns.Add(new DataGridTextColumn
            //{
            //    Header = "Parameters",
            //    Binding = new Binding("Name"),
            //    IsReadOnly = true
            //});

            //// 动态添加 Min/Type/Max 组
            //if (vm.ParameterRows.Count > 0)
            //{
            //    int groupCount = vm.ParameterRows[0].Groups.Count; // 假设所有行组数一致

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

            //// 最后一列：单位
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

            //// 第一列：参数名称
            //featureDataGrid.Columns.Add(new DataGridTextColumn
            //{
            //    Header = "Feature",
            //    Binding = new Binding("Name"),
            //    //IsReadOnly = true
            //});


            //// 最后一列：单位
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
        // 在选中行下方插入
        private void InsertRowBelow_Click(object sender, RoutedEventArgs e)
        {
            InsertRowAtOffset(1); // 插入到当前索引 + 0（即下方）
        }
        // 通用插入方法
        private void InsertRowAtOffset(int offset)
        {
            //if (featureDataGrid.SelectedItems.Count == 0)
            //    return;

            //// 获取第一个选中项（支持多选，但只取第一个）
            //var selectedItem = featureDataGrid.SelectedItems[0];

            //// 找到它在源集合中的索引
            //int currentIndex = vm.FeatureParameterRows.IndexOf(selectedItem as FeatureParameterRow);
            //if (currentIndex == -1)
            //    return;

            //// 计算插入位置
            //int insertIndex = currentIndex + offset;

            //// 边界处理：不能小于 0
            //if (insertIndex < 0)
            //    insertIndex = 0;

            //// 插入新行
            //vm.FeatureParameterRows.Insert(insertIndex, new FeatureParameterRow());

            //// 可选：自动选中新插入的行
            //featureDataGrid.SelectedIndex = insertIndex;
            //featureDataGrid.ScrollIntoView(vm.FeatureParameterRows[insertIndex]);
        }
        private void DeleteFeatureRow_Click(object sender, RoutedEventArgs e)
        {
            // 获取当前选中的项（可能是多个）
            //var selectedItems = new List<object>(featureDataGrid.SelectedItems.Cast<object>());

            //// 从后往前删除（避免索引变化问题）
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
            // 获取触发事件的 ListBox
            var listBox = sender as ListBox;

            // 确保有选中项
            if (listBox?.SelectedItem == null)
                return;

            // 获取选中项的内容（Content）
            var selectedItem = listBox.SelectedItem as ListBoxItem;
            if (selectedItem == null)
                return;

            string content = selectedItem.Content?.ToString();

            // 根据内容判断点击了哪个
            if (content == "Paramtets Table")  // 注意：你拼写的是 "Paramtets"，不是 "Parameters"
            {
                // 切换到 Parameters 表格
                ShowParametersTable();
            }
            else if (content == "Feature Table")
            {
                // 切换到 Feature 表格
                ShowFeatureTable();
            }
        }
        private void ShowParametersTable()
        {
            // 例如：显示参数表格，隐藏 feature 表格
            //dataGrid.Visibility = Visibility.Visible;
            //featureDataGrid.Visibility = Visibility.Collapsed;
        }

        private void ShowFeatureTable()
        {
            // 显示 feature 表格，隐藏参数表格
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
            // 确保 Name 和 Info 有默认值（避免 null 显示问题）
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
