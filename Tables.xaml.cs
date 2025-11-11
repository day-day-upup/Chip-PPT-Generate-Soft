using ChipManualGenerationSogt;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2021.Excel.Pivot;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using HarfBuzzSharp;
using ScottPlot;
using ScottPlot.Colormaps;
using ScottPlot.WPF;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
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
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace ChipManualGenerationSogt
{



    public enum TableType
    {
        Basic,
        Feature,
        Parameters,
        AbsoluteRatings,
        SupplyCurrent,
        Notes,
        Description1,
        Description2,
        TurnOn,
        TurnOff
    }

    class MinMaxType 
    {
        public string Min { get; set; }
        public string Type { get; set; }
        public string Max { get; set; }
    }
    
    
    /// <summary>
    /// Tables.xaml 的交互逻辑
    /// </summary>
    public partial class Tables : UserControl
    {
        TablesModel vm;
        private int _currentGroupCount = 1;
        public Tables()
        {
            InitializeComponent();
            vm = new TablesModel();
            DataContext = vm;
            _currentGroupCount = Properties.Settings.Default.GroupCount;
            if (_currentGroupCount <= 0) _currentGroupCount = 1;
            //InitializeParameterRows(_currentGroupCount);
            //RebuildDataGridColumns();
            //InitializeFeatureParameterRows();
            //FeatureRebuildDataGridColumns();

            //InitializeBasicParameterRows();
            //BasicRebuildDataGridColumns();

            InitializeAllTables();
            RebuildAllDataGridColumns();
            double[] sin = Generate.Sin(51);

            // add a signal plot to the plot
            WpfPlot1.Plot.Add.Signal(sin);
            WpfPlot1.Plot.Title($"Chart");
            WpfPlot1.Plot.YLabel("Voltage (V)");
            WpfPlot1.Plot.XLabel("Time (s)");
            WpfPlot1.Menu.Clear();
            WpfPlot1.Refresh();

        }

        public TablesModel GetModel()
        {
            return vm;        
        }
        public List<string> GetBasicTableInfo()
        {
            List<string> value = new List<string>();
            foreach (var item in vm.BasicParameterRows)
            { 
               string temp = item.Name + ": " + item.Info;
               value.Add(temp);
            }
            return value;

        }


        public string GetFeatureTableInfo()
        {
            string value = "";
            foreach (var item in vm.FeatureParameterRows)
            {
                value += item.Name + ": " + item.Info +"\n";
                
            }
            return value;
        }

        public string[,] GetParameterTableInfo()
        {
            int coulumnCount = 3 * _currentGroupCount + 2;
            string[,] value = new string[vm.ParameterRows.Count+1, coulumnCount];

            value[0, 0] = "Parameter";
            value[0, coulumnCount - 1] = "Unit";
            for (int i = 1; i < coulumnCount - 1; i++)
            {
                if (i % 3 == 1)
                {
                    value[0, i] =  "Min";
                }
                else if (i % 3 == 2)
                {
                    value[0, i] = "Type";
                }
                else
                {
                    value[0, i] = "Max";
                }
            }
            for (int i = 1; i < vm.ParameterRows.Count+1; i++)
            {
                value[i, 0] = vm.ParameterRows[i-1].Name;
                value[i, coulumnCount-1] = vm.ParameterRows[i-1].Unit;
                for (int k = 1; k < coulumnCount-1; k++)
                {
                    if (k % 3 == 1)
                    {
                        value[i, k] = vm.ParameterRows[i-1].Groups[k/3].Min;
                    }
                    else if (k % 3 == 2)
                    {
                        value[i, k] = vm.ParameterRows[i-1].Groups[k/3].Type;
                    }
                    else
                    {
                        value[i, k] = vm.ParameterRows[i-1].Groups[k / 3 -1].Max;
                    }
                    
                
                }
                 
            }
            return value;
        }
        public void PrintParameterTable(string[,] value)
        {
            int rows = value.GetLength(0);
            int cols = value.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    Console.Write($"{value[i, j],-15}"); // 每列宽度固定，左对齐
                }
                Console.WriteLine();
            }
        }

        private void InitializeParameterRows(int groupCount)
        {
            vm.ParameterRows.Clear();

            List<MinMaxType> temp = new List<MinMaxType>();
            var item10= new MinMaxType()
            { 
               Type= "45-70"
            };
            var item11 = new MinMaxType()
            {
               Type = "70-90"
            };
            temp.Add(item10 );
            temp.Add(item11);
            
            AddParameterRow("Frequency", "GHz", groupCount, GenerateGroups(groupCount, temp));

            temp.Clear();
            var item20 = new MinMaxType()
            {
                Min = "14",
                Type = "14.5"
            };
            var item21 = new MinMaxType()
            {
                Min = "16",
                Type = "18"
            };
            temp.Add(item20);
            temp.Add(item21);
            AddParameterRow("Small Signal Gain", "dB", groupCount, GenerateGroups(groupCount, temp));


            temp.Clear();
            var item30 = new MinMaxType()
            {
                Type = "±1.0"
            };
            var item31 = new MinMaxType()
            {
                Type = "±1.0"
            };
            temp.Add(item30);
            temp.Add(item31);
            AddParameterRow("Gain Flatness", "dB", groupCount, GenerateGroups(groupCount, temp));


            temp.Clear();
            var item40 = new MinMaxType()
            {
                Type = "4.5"
            };
            var item41 = new MinMaxType()
            {
                Type = "6.5"
            };
            temp.Add(item40);
            temp.Add(item41);
            AddParameterRow("Noise Figure", "dB", groupCount, GenerateGroups(groupCount, temp));



            temp.Clear();
            var item50 = new MinMaxType()
            {
                Type = "12"
            };
            var item51 = new MinMaxType()
            {
                Type = "14"
            };
            temp.Add(item50);
            temp.Add(item51);
            AddParameterRow("P1dB - Output 1dB Compression", "dBm", groupCount, GenerateGroups(groupCount, temp));

            temp.Clear();
            var item60 = new MinMaxType()
            {
                Type = "15"
            };
            var item61 = new MinMaxType()
            {
                Type = "18"
            };
            temp.Add(item60);
            temp.Add(item61);
            AddParameterRow("Psat - Saturated  Output Power", "dBm", groupCount, GenerateGroups(groupCount, temp));


            temp.Clear();
            var item70 = new MinMaxType()
            {
                Type = "20"
            };
            var item71 = new MinMaxType()
            {
                Type = "22"
            };
            temp.Add(item70);
            temp.Add(item71);
            AddParameterRow("OIP3 - Output Third Order Intercept", "dBm", groupCount, GenerateGroups(groupCount, temp));


            temp.Clear();
            var item80 = new MinMaxType()
            {
                Type = "-18"
            };
            var item81 = new MinMaxType()
            {
                Type = "-13"
            };
            temp.Add(item80);
            temp.Add(item81);
            AddParameterRow("Input Return Loss", "dB", groupCount, GenerateGroups(groupCount, temp));


            temp.Clear();
            var item90 = new MinMaxType()
            {
                Type = "-13"
            };
            var item91 = new MinMaxType()
            {
                Type = "-18"
            };
            temp.Add(item90);
            temp.Add(item91);
            AddParameterRow("Output Return Loss", "dB", groupCount, GenerateGroups(groupCount, temp));


            // 示例：添加两个参数行
            //AddParameterRow("Frequency", "GHz", groupCount, groups);
            //AddParameterRow("Small Signal Gain", "dB", groupCount, groups);
            //AddParameterRow("Gain Flatness", "dB", groupCount, groups);
            //AddParameterRow("Noise Figure", "dB", groupCount, groups);
            //AddParameterRow("P1dB - Output 1dB Compression", "dBm", groupCount, groups);
            //AddParameterRow("Psat - Saturated  Output Power", "dBm", groupCount, groups);
            //AddParameterRow("OIP3 - Output Third Order Intercept", "dBm", groupCount, groups);
            //AddParameterRow("Input Return Loss", "dB", groupCount, groups);
            //AddParameterRow("Output Return Loss", "dB", groupCount, groups);
        }
        private ObservableCollection<MinMaxTypeGroup> GenerateGroups(int groupCount , List<MinMaxType> value)
        { 
            var groups = new ObservableCollection<MinMaxTypeGroup>();
            List<MinMaxTypeGroup> groupList = new List<MinMaxTypeGroup>();
            
            for (int i = 0; i < groupCount; i++)
            {

                var item = new MinMaxTypeGroup()
                {
                   Min = value.ElementAt(i).Min,
                   Type = value.ElementAt(i).Type,
                   Max = value.ElementAt(i).Max
                };
                groupList.Add(item);
            }

            foreach (var item in groupList) 
            {
                groups.Add(item);
            }
            return groups;
        }
        private void Clearn(MinMaxTypeGroup group) 
        {
            if (group == null) return;

            group.Min = string.Empty; 
            group.Type = string.Empty; 
            group.Max = string.Empty; 
        }
        private void InitializeFeatureParameterRows()
        {
            vm.FeatureParameterRows.Clear();
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

        private void InitializeBasicParameterRows()
        {
            vm.BasicParameterRows.Clear();
            BasicAddParameterRow("Manual PN", "MML806");
            BasicAddParameterRow("Version", "V3.0.0");
            BasicAddParameterRow("Product Name", "±2.5dB Typical");
            BasicAddParameterRow("Frequency Range", "45-90GHz");
            BasicAddParameterRow("Right Slider Info", "4.5dB Typical");
        }

        public void SetBasicParameterRow(string str)
        {
            vm.BasicParameterRows.ElementAt(0).Info = str;
        }
        private void AddParameterRow(string name, string unit, int groupCount, ObservableCollection<MinMaxTypeGroup> groups)
        {
            var row = new ParameterRow
            {
                Name = name,
                Unit = unit,
                Groups = new ObservableCollection<MinMaxTypeGroup>()
                //Groups = groups
            };

            for (int i = 0; i < groupCount; i++)
            {
                row.Groups.Add(groups.ElementAt(i));
            }

            vm.ParameterRows.Add(row);
        }

        private void FeatureAddParameterRow(string name, string info)
        {
            var row = new FeatureParameterRow
            {
                Name = name,
                Info = info
            };
            vm.FeatureParameterRows.Add(row);
        }

        private void BasicAddParameterRow(string name, string info)
        {
            var row = new FeatureParameterRow
            {
                Name = name,
                Info = info
            };
            vm.BasicParameterRows.Add(row);
        }

        private void RebuildDataGridColumns()
        {
            parametersDataGrid.Columns.Clear();

            // 第一列：参数名称
            parametersDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Parameters",
                Binding = new Binding("Name"),
                IsReadOnly = true
            });

            // 动态添加 Min/Type/Max 组
            if (vm.ParameterRows.Count > 0)
            {
                int groupCount = vm.ParameterRows[0].Groups.Count; // 假设所有行组数一致

                for (int i = 0; i < groupCount; i++)
                {
                    parametersDataGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = $"Min {i + 1}",
                        Binding = new Binding($"Groups[{i}].Min")
                    });

                    parametersDataGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = $"Type {i + 1}",
                        Binding = new Binding($"Groups[{i}].Type")
                    });

                    parametersDataGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = $"Max {i + 1}",
                        Binding = new Binding($"Groups[{i}].Max")
                    });
                }
            }

            // 最后一列：单位
            parametersDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Units",
                Binding = new Binding("Unit"),
                IsReadOnly = true
            });
        }

        private void FeatureRebuildDataGridColumns()
        {
            featureDataGrid.Columns.Clear();

            // 第一列：参数名称
            featureDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Feature",
                Binding = new Binding("Name"),
                //IsReadOnly = true
            });


            // 最后一列：单位
            featureDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = " ",
                Binding = new Binding("Info"),
                //IsReadOnly = true
            });
        }


        private void BasicRebuildDataGridColumns()
        {
            basicDataGrid.Columns.Clear();

            // 第一列：参数名称
            basicDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Feature",
                Binding = new Binding("Name"),
                //IsReadOnly = true
            });


            // 最后一列：单位
            basicDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = " ",
                Binding = new Binding("Info"),
                //IsReadOnly = true
            });
        }

        private void AddGroup_Click(object sender, RoutedEventArgs e)
        {
            _currentGroupCount++;
            Properties.Settings.Default.GroupCount = _currentGroupCount;
            foreach (var row in vm.ParameterRows)
            {
                row.Groups.Add(new MinMaxTypeGroup());
            }
            RebuildDataGridColumns();
        }

        private void RemoveGroup_Click(object sender, RoutedEventArgs e)
        {
            if (_currentGroupCount > 1)
            {
                _currentGroupCount--;
                foreach (var row in vm.ParameterRows)
                {
                    row.Groups.RemoveAt(row.Groups.Count - 1);
                }
                RebuildDataGridColumns();
                Properties.Settings.Default.GroupCount = _currentGroupCount;
            }
        }

        private void AddFeatureRow_Click(object sender, RoutedEventArgs e)
        {
           
            vm.FeatureParameterRows.Add(new FeatureParameterRow());
        }
        private void InsertRowAbove_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            var contextMenu = menuItem?.Parent as ContextMenu;
            var dataGrid = contextMenu?.PlacementTarget as DataGrid;

            if (dataGrid != null)
            {
                var collection = GetCollectionForDataGrid(dataGrid);
                if (collection != null)
                    InsertRowAtOffset(collection, dataGrid, 0);
            }
        }
        private void InsertRowAtOffset(ObservableCollection<FeatureParameterRow> collection, DataGrid dataGrid, int offset)
        {
            if (dataGrid.SelectedItems.Count == 0)
                return;

            var selectedItem = dataGrid.SelectedItems[0];
            int currentIndex = collection.IndexOf(selectedItem as FeatureParameterRow);
            if (currentIndex == -1)
                return;

            int insertIndex = Math.Max(0, currentIndex + offset);
            collection.Insert(insertIndex, new FeatureParameterRow());

            dataGrid.SelectedIndex = insertIndex;
            dataGrid.ScrollIntoView(collection[insertIndex]);
        }
        // 在选中行下方插入
        private void InsertRowBelow_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            var contextMenu = menuItem?.Parent as ContextMenu;
            var dataGrid = contextMenu?.PlacementTarget as DataGrid;

            if (dataGrid != null)
            {
                var collection = GetCollectionForDataGrid(dataGrid);
                if (collection != null)
                    InsertRowAtOffset(collection, dataGrid, 1);
            }
        }// 通用插入方法
        private void InsertRowAtOffset(int offset)
        {
            if (featureDataGrid.SelectedItems.Count == 0)
                return;

            // 获取第一个选中项（支持多选，但只取第一个）
            var selectedItem = featureDataGrid.SelectedItems[0];

            // 找到它在源集合中的索引
            int currentIndex = vm.FeatureParameterRows.IndexOf(selectedItem as FeatureParameterRow);
            if (currentIndex == -1)
                return;

            // 计算插入位置
            int insertIndex = currentIndex + offset;

            // 边界处理：不能小于 0
            if (insertIndex < 0)
                insertIndex = 0;

            // 插入新行
            vm.FeatureParameterRows.Insert(insertIndex, new FeatureParameterRow());

            // 可选：自动选中新插入的行
            featureDataGrid.SelectedIndex = insertIndex;
            featureDataGrid.ScrollIntoView(vm.FeatureParameterRows[insertIndex]);
        }
        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            var contextMenu = menuItem?.Parent as ContextMenu;
            var dataGrid = contextMenu?.PlacementTarget as DataGrid;

            if (dataGrid != null)
            {
                var collection = GetCollectionForDataGrid(dataGrid);
                if (collection != null)
                    DeleteRows(collection, dataGrid);
            }
        }

        private void DeleteRows(ObservableCollection<FeatureParameterRow> collection, DataGrid dataGrid)
        {
            var selectedItems = new List<object>(dataGrid.SelectedItems.Cast<object>());
            foreach (var item in selectedItems.Reverse<object>())
            {
                if (item is FeatureParameterRow row)
                {
                    collection.Remove(row);
                }
            }
        }
        private ObservableCollection<FeatureParameterRow> GetCollectionForDataGrid(DataGrid dataGrid)
        {
            if (dataGrid == null) return null;

            switch (dataGrid.Name)
            {
                case "basicDataGrid":
                    return vm.BasicParameterRows;
                case "featureDataGrid":
                    return vm.FeatureParameterRows;
                case "absoluteRatingsDataGrid":
                    return vm.AbsoluteRatingsRows;
                case "supplyCurrentDataGrid":
                    return vm.SupplyCurrentRows;
                case "notesDataGrid":
                    return vm.NotesRows;
                case "description1DataGrid":
                    return vm.Description1Rows;
                case "description2DataGrid":
                    return vm.Description2Rows;
                case "turnOnDataGrid":
                    return vm.TurnOnRows;
                case "turnOffDataGrid":
                    return vm.TurnOffRows;
                default:
                    return null;
            }
        }

        private void ShowTable(TableType tableType)
        {
            // 隐藏所有内容
            HideAllContent();

            // 根据选择的表格类型显示相应内容
            switch (tableType)
            {
                case TableType.Basic:
                    {
                        basicDataGrid.Visibility = Visibility.Visible;
                        basicTableImage.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.Feature:
                    {
                        featureDataGrid.Visibility = Visibility.Visible;
                        featureImage.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.Parameters:
                    {
                        parametersDataGrid.Visibility = Visibility.Visible;
                        elTableImage.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.AbsoluteRatings:
                    {
                        absoluteRatingsImage.Visibility = Visibility.Visible;
                        absoluteRatingsDataGrid.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.SupplyCurrent:
                    {
                        supplyCurrentDataGrid.Visibility = Visibility.Visible;
                        supplyCurrentImage.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.Notes:
                    {
                        notesDataGrid.Visibility = Visibility.Visible;
                        notesImage.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.Description1:
                    {
                        description1DataGrid.Visibility = Visibility.Visible;
                        description1Image.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.Description2:
                    {
                        description2DataGrid.Visibility = Visibility.Visible;
                        description2Image.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.TurnOn:
                    {
                        turnOnDataGrid.Visibility = Visibility.Visible;
                        turnOnImage.Visibility = Visibility.Visible;
                    }
                    break;

                case TableType.TurnOff:
                    {
                        turnOffDataGrid.Visibility = Visibility.Visible;
                        turnOffImage.Visibility = Visibility.Visible;
                    }
                    break;
                
            }
        }

        // 定义表格类型枚举

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

            // 根据内容转换为 TableType
            TableType tableType = GetTableTypeFromContent(content);

            // 显示对应的表格
            ShowTable(tableType);
        }

        private TableType GetTableTypeFromContent(string content)
        {
            if (string.IsNullOrEmpty(content))
                return TableType.Basic;

            switch (content)
            {
                case "Basic Table":
                    return TableType.Basic;
                case "Feature Table":
                    return TableType.Feature;
                case "Paramtets Table":
                    return TableType.Parameters;
                case "Absolute Maximum Ratings":
                    return TableType.AbsoluteRatings;
                case "Typical Supply Current vs. VD,VG":
                    return TableType.SupplyCurrent;
                case "Notes":
                    return TableType.Notes;
                case "Description1":
                    return TableType.Description1;
                case "Description2":
                    return TableType.Description2;
                case "Turn ON procedure":
                    return TableType.TurnOn;
                case "Turn OFF procedure":
                    return TableType.TurnOff;
                default:
                    return TableType.Basic;
            }
        }


        private void HideAllContent()
        {
            // 隐藏所有表格
            parametersDataGrid.Visibility = Visibility.Collapsed;
            featureDataGrid.Visibility = Visibility.Collapsed;
            basicDataGrid.Visibility = Visibility.Collapsed;
            absoluteRatingsDataGrid.Visibility = Visibility.Collapsed;
            supplyCurrentDataGrid.Visibility = Visibility.Collapsed;
            notesDataGrid.Visibility = Visibility.Collapsed;
            description1DataGrid.Visibility = Visibility.Collapsed;
            description2DataGrid.Visibility = Visibility.Collapsed;
            turnOffDataGrid.Visibility = Visibility.Collapsed;
            turnOnDataGrid.Visibility = Visibility.Collapsed;


            // 隐藏所有图片内容
            basicTableImage.Visibility = Visibility.Collapsed;
            featureImage.Visibility = Visibility.Collapsed;
            elTableImage.Visibility = Visibility.Collapsed;
            absoluteRatingsImage.Visibility = Visibility.Collapsed;
            supplyCurrentImage.Visibility = Visibility.Collapsed;
            notesImage.Visibility = Visibility.Collapsed;
            description1Image.Visibility = Visibility.Collapsed;
            description2Image.Visibility = Visibility.Collapsed;
            turnOnImage.Visibility = Visibility.Collapsed;
            turnOffImage.Visibility = Visibility.Collapsed;

            // 隐藏图表
            WpfPlot1.Visibility = Visibility.Collapsed;
        }

        private void InitializeAllTables()
        {
            // 参数表（特殊处理）
            InitializeParameterRows(_currentGroupCount);
            RebuildDataGridColumns();

            // 其他表格使用统一的初始化方法
            InitializeTableData(vm.BasicParameterRows, GetBasicTableData());
            InitializeTableData(vm.FeatureParameterRows, GetFeatureTableData());
            InitializeTableData(vm.AbsoluteRatingsRows, GetAbsoluteRatingsData());
            InitializeTableData(vm.SupplyCurrentRows, GetSupplyCurrentData());
            InitializeTableData(vm.NotesRows, GetNotesData());
            InitializeTableData(vm.Description1Rows, GetDescription1Data());
            InitializeTableData(vm.Description2Rows, GetDescription2Data());
            InitializeTableData(vm.TurnOnRows, GetTurnOnData());
            InitializeTableData(vm.TurnOffRows, GetTurnOffData());

            // 为所有表格重建列（除了参数表）
            RebuildAllDataGridColumns();
        }

        // 统一的表格数据初始化方法
        private void InitializeTableData(ObservableCollection<FeatureParameterRow> collection, List<(string name, string info)> data)
        {
            collection.Clear();
            foreach (var (name, info) in data)
            {
                collection.Add(new FeatureParameterRow { Name = name, Info = info });
            }
        }

        // 各个表格的数据定义
        private List<(string name, string info)> GetBasicTableData()
        {
            return new List<(string, string)>
    {
        ("Manual PN", "MML806"),
        ("Version", "V3.0.0"),
        ("Product Name", "Your Product Name"),
        ("Frequency Range", "45-90GHz"),
        ("Right Slider Info", "Additional Information")
    };
        }

        private List<(string name, string info)> GetFeatureTableData()
        {
            return new List<(string, string)>
    {
        ("Frequency", "45-90GHz"),
        ("Small Signal Gain", "15dB Typical"),
        ("Gain Flatness", "±2.5dB Typical"),
        ("Noise Figure", "4.5dB Typical"),
        ("P1dB", "12dBm Typical"),
        ("Power Supply", "VD=+4V@119mA ,VG=-0.4V"),
        ("Input/Output", "50Ω"),
        ("Chip Size", "1.766 x 2.0 x 0.05mm")
    };
        }

        private List<(string name, string info)> GetAbsoluteRatingsData()
        {
            return new List<(string, string)>
    {
        ("Storage Temperature", "-65 to +150 °C"),
        ("Operating Temperature", "-55 to +125 °C"),
        ("Supply Voltage VD", "-0.5 to +5 V"),
        ("Supply Voltage VG", "-2 to +0.5 V"),
        ("RF Input Power", "+15 dBm")
    };
        }

        private List<(string name, string info)> GetSupplyCurrentData()
        {
            return new List<(string, string)>
    {
        ("VD = +3V", "85 mA"),
        ("VD = +4V", "119 mA"),
        ("VD = +5V", "150 mA")
    };
        }

        private List<(string name, string info)> GetNotesData()
        {
            return new List<(string, string)>
    {
        ("Note 1", "Important note content here"),
        ("Note 2", "Another important note")
    };
        }

        private List<(string name, string info)> GetDescription1Data()
        {
            return new List<(string, string)>
    {
        ("Description", "First description content")
    };
        }

        private List<(string name, string info)> GetDescription2Data()
        {
            return new List<(string, string)>
    {
        ("Description", "Second description content")
    };
        }

        private List<(string name, string info)> GetTurnOnData()
        {
            return new List<(string, string)>
    {
        ("Step 1", "Apply VG voltage"),
        ("Step 2", "Apply VD voltage"),
        ("Step 3", "Verify operation")
    };
        }

        private List<(string name, string info)> GetTurnOffData()
        {
            return new List<(string, string)>
    {
        ("Step 1", "Remove VD voltage"),
        ("Step 2", "Remove VG voltage")
    };
        }

        private void RebuildDataGridColumns(DataGrid dataGrid, string firstColumnHeader = "Feature", string secondColumnHeader = " ")
        {
            dataGrid.Columns.Clear();

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = firstColumnHeader,
                Binding = new Binding("Name")
            });

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = secondColumnHeader,
                Binding = new Binding("Info")
            });
        }

        private void RebuildAllDataGridColumns()
        {
            // 参数表使用特殊列重建
            RebuildDataGridColumns();

            // 其他表格使用统一列重建
            RebuildDataGridColumns(basicDataGrid, "Basic", "Information");
            RebuildDataGridColumns(featureDataGrid, "Feature", "Value");
            RebuildDataGridColumns(absoluteRatingsDataGrid, "Parameter", "Rating");
            RebuildDataGridColumns(supplyCurrentDataGrid, "Condition", "Current");
            RebuildDataGridColumns(notesDataGrid, "Note", "Content");
            RebuildDataGridColumns(description1DataGrid, "Item", "Description");
            RebuildDataGridColumns(description2DataGrid, "Item", "Description");
            RebuildDataGridColumns(turnOnDataGrid, "Step", "Procedure");
            RebuildDataGridColumns(turnOffDataGrid, "Step", "Procedure");
        }
    }

    public  class TablesModel : ObeservableObject
    {
        public ObservableCollection<ParameterRow> ParameterRows { get; set; } = new ObservableCollection<ParameterRow>();
        public ObservableCollection<FeatureParameterRow> FeatureParameterRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> BasicParameterRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> AbsoluteRatingsRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> SupplyCurrentRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> NotesRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> Description1Rows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> Description2Rows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> TurnOnRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> TurnOffRows { get; set; } = new ObservableCollection<FeatureParameterRow>();

    }


    public class MinMaxTypeGroup : INotifyPropertyChanged
    {
        private string _min;
        private string _type;
        private string _max;

        public string Min
        {
            get => _min;
            set { _min = value; OnPropertyChanged(); }
        }

        public string Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(); }
        }

        public string  Max
        {
            get => _max;
            set { _max = value; OnPropertyChanged(); }
        }

        public MinMaxTypeGroup() 
        {
        }
        public MinMaxTypeGroup(string min, string type, string max)
        {
            Min = min;
            Type = type;
            Max = max;
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

           