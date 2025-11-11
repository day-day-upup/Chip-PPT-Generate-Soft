using ChipManualGenerationSogt;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2021.Excel.Pivot;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using HarfBuzzSharp;
using OpenTK.Graphics.ES11;
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
using System.Windows.Markup;
using System.Text.Json;
using Microsoft.Office.Interop.PowerPoint;
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
        //List<(string name, string info, string info2)> myData;
        TablesModel vm;
        private int _currentGroupCount = 1;
        public Tables()
        {
            InitializeComponent();
            vm = new TablesModel();
            DataContext = vm;
            //_currentGroupCount = Properties.Settings.Default.GroupCount;
            _currentGroupCount = 2;
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
                //string temp = item.Name + ": " + item.Info;
                string temp =item.Info;
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
                    value[0, i] =  "Min.";
                }
                else if (i % 3 == 2)
                {
                    value[0, i] = "Type.";
                }
                else
                {
                    value[0, i] = "Max.";
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


            DataTemplate showMinTemplate = (DataTemplate)this.FindResource("ShowMinTemplate");
            parametersDataGrid.Columns.Add(new DataGridTemplateColumn
            {
                Header = "Show Min",
                CellTemplate = showMinTemplate,
                //Width = 30
            });
            DataTemplate showMaxTemplate = (DataTemplate)this.FindResource("ShowMaxTemplate");
            // ? 新增：Show Max CheckBox 列
            parametersDataGrid.Columns.Add(new DataGridTemplateColumn
            {
                Header = "Show Max",
                CellTemplate = showMaxTemplate,
                //Width =20
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
           // Properties.Settings.Default.GroupCount = _currentGroupCount;
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
                //Properties.Settings.Default.GroupCount = _currentGroupCount;
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
                    //return vm.SupplyCurrentRows;
                case "notesDataGrid":
                    //return vm.NotesRows;
                case "description1DataGrid":
                    return vm.Description1Rows;
                case "description2DataGrid":
                    //return vm.Description2Rows;
                case "turnOnDataGrid":
                   // return vm.TurnOnRows;
                case "turnOffDataGrid":
                    //return vm.TurnOffRows;
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
            var basicTable = new List<(string, string)>
                            {
                                ("{Top PN}", "MML806"),
                                ("{Version}", "V1.0.0"),
                                ("{Product Name}", "GaAs MMIC Low Noise Amplifier"),
                                ("{Frequency Range}", "0-10GHz"),
                                ("{right bar info}", "GaAs Low Noise Amplifier MMIC 45 – 90GHz")
                            }; 
            InitializeTableData(vm.BasicParameterRows, basicTable);
            var featureTable =new List<(string, string)>
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

            InitializeTableData(vm.FeatureParameterRows, featureTable);

            var rateTable = new List<(string, string)>
            {
                ("Drain Bias Voltage (VD)", "+4.5V"),
        ("Gate Bias Voltage (VG)", "-2V to 0V"),
        ("RF Input Power (RFIN)", "+15dBm"),
        ("Continuous Pdiss (T = 85 °C)\n (derate 6.1mW/°C above 85 °C) ", "175°C"),
        ("Thermal Resistance\n (channel to die bottom)", "0.55W"),
        ("Operating Temperature", "-55°C to +85 °C"),
        ("Storage Temperature", "-65°C to +150 °C"),
    }
            ;

            InitializeTableData(vm.AbsoluteRatingsRows, rateTable);


           var supplyCurrentTable = new List<(string, string,string)>
    {
       
        ("+3.5", "-0.38","118"),
        ("4.0", "-0.40","119"),
        ("4.0", "-0.50","71")
    };
            InitializeTable(vm.CurrentVdVgTable, supplyCurrentTable);


            var notesTable = new List<string>
    {
        ( "1. Die thickness: 50μm"),
        ( "2. VD bond pad is 75*75μm?"),
        ( "3. VG bond pad is 75*75μm?"),
        ( "4. RF IN/OUT bond pad is 50*86μm?"),
        ( "5. Bond pad metalization: Gold"),
        ( "6. Backside metalization: Gold"),
    };

            InitializeTable(vm.NotesTable, notesTable);

            var d1Table =new List<(string, string)>
    {
        ("C1", "100pF Example: Skyworks Part: SC10002430"),
        ("C2", "0.01μF Example: TDK Part:C1005X7R1H103K050BB (0402)"),
        ("C3", "0.1μF Example: TDK Part:C1005X7R1H104K050BB (0402)"),
        ("R1", "100Ω Example: Yageo Part:SR0402FR-7T10RL"),
    };
            InitializeTableData(vm.Description1Rows, d1Table);

            var d2Table = new List<(string, string,string)>
    {
        ("1", "RF IN","RF signal input terminal; no blocking capacitor required."),
        ("2", "RF OUT","RF signal output terminal; no blocking capacitor required."),
        ("3", "VD","Drain Biases for the Amplifier ; An external biasing circuit is required."),
        ("4", "VG","Gate Biases for the Amplifier ; An external biasing circuit is required."),
        ("5", "Die Bottom","Die bottom must be connected to RF and dc ground.")
    };
            InitializeTable(vm.Description2Rows, d2Table);

           

            var t1Table = new List<string>
                    {
                        "Turn ON procedure:",
                        "1.    Connect GND to RF and dc ground.",
                        "2.    Set the gate bias voltages VG to -2V.",
                        "3.    Set the drain bias voltages VD to +4V.",
                        "4.    Increase the gate bias voltages to achieve a quiescent supply current of 82 mA.",
                        "5.    Apply RF signal.",
                    };
            InitializeTable(vm.TurnOnRows, t1Table);


            var t2Table = new List<string>
                    {
                        "Turn OFF procedure:",
                        "1.    Turn off the RF signal.",
                        "2.    Decrease the gate bias voltages, VG to -2V to achieve a IDQ = 0 mA (approximately).",
                        "3.    Decrease the drain bias voltages to 0 V.",
                        "4.    Increase the all gate bias voltages to 0 V.",
                    };
            InitializeTable(vm.TurnOffRows, t2Table);

            // 为所有表格重建列（除了参数表）
            RebuildAllDataGridColumns();
        }

        // 统一的表格数据初始化方法
        private void InitializeTableDataThree(ObservableCollection<TreeColumnTableRow> collection, List<(string first, string send, string third)> data)
        {
            collection.Clear();
            foreach (var (first, send, third) in data)
            {
                collection.Add(new TreeColumnTableRow{ FirstColumn = first, SecondColumn = send, ThirdColumn= third });
            }
        }

        private void InitializeTableData(ObservableCollection<FeatureParameterRow> collection, List<(string first, string send)> data)
        {
            collection.Clear();
            foreach (var (first, send) in data)
            {
                collection.Add(new FeatureParameterRow { Name= first, Info = send });
            }
        }
        /// <summary>
        /// 处理三列的表格数据初始化
        /// </summary>
        /// <param name="collection"></param>
        /// <param name="data"></param>
        private void InitializeTable(ObservableCollection<TreeColumnTableRow> collection, List<(string first, string send, string third)> data)
        {
            collection.Clear();
            foreach (var (first, send, third) in data)
            {
                collection.Add(new TreeColumnTableRow { FirstColumn= first, SecondColumn =send, ThirdColumn= third });
            }

        }

        private void InitializeTable(ObservableCollection<OneColumnTableRow> collection, List<string > data)
        {
            collection.Clear();
            foreach (var item in data)
            {
                collection.Add(new OneColumnTableRow { FirstColumn = item });
            }

        }
        //private void InitializeTableDataOne(ObservableCollection<TreeColumnTableRow> collection, List<(string vd, string vg, string idd)> data)
        //{
        //    collection.Clear();
        //    foreach (var (name, info) in data)
        //    {
        //        collection.Add(new FeatureParameterRow { Name = name, Info = info });
        //    }
        //}


        private List<(string, string)> GetTwoMetaTable(ObservableCollection<FeatureParameterRow> table)
        { 
            var resoult = new List<(string, string)>();
            foreach (var row in table)
            {
                (string, string) entry = (row.Name, row.Info);
                resoult.Add(entry);
            }
            return resoult;
        }
        // 各个表格的数据定义
        public List<(string name, string info)> GetBasicTableData()
        {
            return GetTwoMetaTable(vm.BasicParameterRows);
        }
        
        public List<(string name, string info)> GetFeatureTableData()
        {
            return GetTwoMetaTable(vm.FeatureParameterRows);
        }

        /// <summary>
        /// 参数表格特殊， 返回值要特殊处理
        /// </summary>
        /// <returns></returns>
        public List<(string name, List<string> value, string unit)> GetParametersTableData()
        {

            var resoult = new List<(string name, List<string> value, string unit)>();

            int valueColunmCount = 3 * _currentGroupCount;
            List<string> headerValue = new List<string>();
            for (int i = 0; i < valueColunmCount - 1; i++)
            {
                if (i % 3 == 1)
                {
                    headerValue.Add("Min");
                }
                else if (i % 3 == 2)
                {
                    //value[0, i] = "Type";
                    headerValue.Add("Type");
                }
                else
                {
                    //value[0, i] = "Max";
                    headerValue.Add("Max");
                }
            }
            (string, List<string>, string) header = ("Parameter", headerValue, "Unit");
            resoult.Add(header);

            foreach (var item in vm.ParameterRows)
            {
                List<string> tmpValue = new List<string>();
                foreach (var group in item.Groups)
                {
                    tmpValue.Add(group.Min);
                    tmpValue.Add(group.Type);
                    tmpValue.Add(group.Max);
                }
                (string, List<string>, string) entry = (item.Name, tmpValue, item.Unit);
                resoult.Add(entry);
            }
            
            //int coulumnCount = 3 * _currentGroupCount + 2;
            //string[,] value = new string[vm.ParameterRows.Count + 1, coulumnCount];
            
            //for (int i = 1; i < vm.ParameterRows.Count + 1; i++)
            //{
            //    value[i, 0] = vm.ParameterRows[i - 1].Name;
            //    value[i, coulumnCount - 1] = vm.ParameterRows[i - 1].Unit;
            //    for (int k = 1; k < coulumnCount - 1; k++)
            //    {
            //        if (k % 3 == 1)
            //        {
            //            value[i, k] = vm.ParameterRows[i - 1].Groups[k / 3].Min;
            //        }
            //        else if (k % 3 == 2)
            //        {
            //            value[i, k] = vm.ParameterRows[i - 1].Groups[k / 3].Type;
            //        }
            //        else
            //        {
            //            value[i, k] = vm.ParameterRows[i - 1].Groups[k / 3 - 1].Max;
            //        }


            //    }

            //}
            
            
            return resoult;

        }
        public List<(string name, string info)> GetAbsoluteRatingsData()
        {
            return GetTwoMetaTable(vm.AbsoluteRatingsRows);
        }

        /// <summary>
        /// 电压， 电流表
        /// </summary>
        /// <returns></returns>
        public List<(string vd, string vg, string idd )> GetCurrentVdVgData()
        {
            var resoult = new List<(string, string, string)>();
             (string, string, string) tem = ("VD(V)", "VG(V)", "IDQ(mA)" );
            resoult.Add(tem);
            foreach (var item in vm.CurrentVdVgTable)
            {
                var tmp = (item.FirstColumn, item.SecondColumn, item.ThirdColumn);
                resoult.Add(tmp);
            }

            return resoult; 
        }

        public List<string> GetNotesData()
        {
            var resoult = new List<string>();
            resoult.Add("Notes:");
            foreach (var item in vm.NotesTable)
            {
               
                resoult.Add(item.FirstColumn);
            }

            return resoult;
        }

        public List<(string name, string info)> GetDescription1Data()
        {
            return GetTwoMetaTable(vm.Description1Rows);
        }

        public List<(string firstColumn, string secondColumn, string thirdColumn)> GetDescription2Data()
        {
            var resoult = new List<(string, string, string)>();
            foreach (var item in vm.Description2Rows)
            {
                var tmp = (item.FirstColumn, item.SecondColumn, item.ThirdColumn);
                resoult.Add(tmp);
            }

            return resoult;
        }

        public List<string> GetTurnOnData()
        {
            var resoult = new List<string>();
            foreach (var item in vm.TurnOnRows)
            {
                resoult.Add(item.FirstColumn);
            }
            return resoult;
        }

        public List<string> GetTurnOffData()
        {
            var resoult = new List<string>();
            foreach (var item in vm.TurnOffRows)
            {
                resoult.Add(item.FirstColumn);
            }
            return resoult;
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

        private void RebuildDataGridColumns(DataGrid dataGrid, string firstColumnHeader = "Feature")
        {
            dataGrid.Columns.Clear();
            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = firstColumnHeader,
                Binding = new Binding("FirstColumn")
            });
        }

        private void RebuildDataGridColumns(DataGrid dataGrid, string firstColumnHeader = "Feature", string SecondColumnHeader = "Feature", string thirdColumnHeader = "Feature")
        {
            dataGrid.Columns.Clear();
            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = firstColumnHeader,
                Binding = new Binding("FirstColumn")
            });
            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = SecondColumnHeader,
                Binding = new Binding("SecondColumn")
            });
            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = thirdColumnHeader,
                Binding = new Binding("ThirdColumn")
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
            RebuildDataGridColumns(supplyCurrentDataGrid, "VD(V)", "VG(V)","IDD(mA)");
            RebuildDataGridColumns(notesDataGrid, "Content");
            RebuildDataGridColumns(description1DataGrid, "Item", "Description");
            RebuildDataGridColumns(description2DataGrid, "No", "Function","Description");
            RebuildDataGridColumns(turnOnDataGrid,  "Content");
            RebuildDataGridColumns(turnOffDataGrid,  "Content");
        }

        private void SetBasicTableInfo() 
        {
        
        
        }
        public void SetParametersTableInfo(List<List<double>> data)
        {
            //vm.ParameterRows[1].IsMinVisible = false;
            //vm.ParameterRows[1].IsMinVisible = true;
            // 假设 vm 和 _currentGroupCount 已经被定义

            // 1. 外部数据完整性检查
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }
            if (data.Count != vm.ParameterRows.Count)
            {
                throw new Exception("The parameter table data rows count is not right.");
            }
            // 确保数据至少有一行，以便检查列数
            if (data.Count > 0 && data[0].Count != 3 * _currentGroupCount)
            {
                throw new Exception("The parameter table data columns count is not right.");
            }

            // 2. 遍历数据行 (i)
            for (int i = 0; i < data.Count; i++)
            {
                // 3. 遍历数据列 (j)。使用标准的 for 循环并递增 j += 3
                // j 代表当前组的起始列索引 (Min, Type, Max)
                for (int j = 0; j < data[i].Count; j += 3)
                {
  
                    var targetGroup = vm.ParameterRows[i].Groups[j / 3];
                    double valMin = data[i][j];      // Min 值
                    double valType = data[i][j + 1]; // Type/Nominal 值
                    double valMax = data[i][j + 2];  // Max 值
                    if (i == 0)
                    {

                        targetGroup.Min = "";
                        targetGroup.Max = "";
                        targetGroup.Type = valMin.ToString("f0") + " - " + valMax.ToString("f0");
                    }
                    // i=2 的特殊逻辑：Type 设置为 ±值
                    else if (i == 2)
                    {
                        
                        string tmp = data[i][j+1].ToString("f1");
                        targetGroup.Min = "";
                        targetGroup.Type = "±" + tmp; 
                        targetGroup.Max = "";
                    }
                    // 其他行 i!=0 且 i!=2 的通用逻辑
                    else
                    {
                        
                        //double valMin = data[i][j];      // Min 值
                        //double valType = data[i][j + 1]; // Type/Nominal 值
                        //double valMax = data[i][j + 2];  // Max 值

                        
                        //targetGroup.Min = valMin == 0 ? "" : valMin.ToString("f1");

                        
                        targetGroup.Type = valType == 0 ? "" : valType.ToString("f1");

                        
                       // targetGroup.Max = valMax == 0 ? "" : valMax.ToString("f1");


                        targetGroup.SetRawValues(valMin.ToString("f1"), valMax.ToString("f1"));
                    }
                }
            }
        }


        /// <summary>
        /// 将 FeatureParameterRow 集合转换为 JSON 模型所需的 ParameterItem 集合 (string, string) => (Key, Value)。
        /// </summary>
        private List<ParameterItem> MapFeatureRowsToParameterItems(
            IEnumerable<FeatureParameterRow> rows)
        {
            // FeatureParameterRow 包含 Name 和 Info 属性，对应 JSON 模型的 Key 和 Value
            return rows.Select(r => new ParameterItem
            {
                Key = r.Name,
                Value = r.Info
            }).ToList();
        }


        public ProductData CreateProductDataFromViewModel()
        {
            // 注意：您代码中 ModelNumber 和 ProductName 的值没有明确从 vm 属性中获取，
            // 我们在此假设它们在 ViewModel 上有相应的属性，或者您需要手动提供它们。

            var productData = new ProductData
            {
                // 假设 ViewModel 上有 ProductName 和 ModelNumber 属性
                ProductName = "GaAs MMIC Low Noise Amplifier", // 替换为 vm.ProductName
                ModelNumber = "MML806", // 替换为 vm.ModelNumber

                Tables = new ParamTables
                {
                    // ------------------ 映射两列数据表 ------------------
                    // BasicParameters, FeatureParameters, AbsoluteMaximumRatings, Description1 结构相同
                    BasicParameters = MapFeatureRowsToParameterItems(vm.BasicParameterRows),
                    FeatureParameters = MapFeatureRowsToParameterItems(vm.FeatureParameterRows),
                    AbsoluteMaximumRatings = MapFeatureRowsToParameterItems(vm.AbsoluteRatingsRows),
                    Description1 = vm.Description1Rows.Select(r => new DescriptionItem1
                    {
                        Component = r.Name,
                        Description = r.Info // FeatureParameterRow 的 Info 属性对应 Description
                    }).ToList(),

                    // ------------------ 映射三列数据表 ------------------
                    SupplyCurrentVdVg = vm.CurrentVdVgTable.Select(r => new SupplyCurrentItem
                    {
                        VD = r.FirstColumn,
                        VG = r.SecondColumn,
                        Current_mA = r.ThirdColumn // 记住，Current (mA) 映射到 Current_mA
                    }).ToList(),

                    Description2 = vm.Description2Rows.Select(r => new DescriptionItem2
                    {
                        Pin = r.FirstColumn,
                        Function = r.SecondColumn,
                        Detail = r.ThirdColumn
                    }).ToList(),

                    // ------------------ 映射单列数据表 ------------------
                    Notes = vm.NotesTable.Select(r => r.FirstColumn).ToList(),
                    TurnOnProcedure = vm.TurnOnRows.Select(r => r.FirstColumn).ToList(),
                    TurnOffProcedure = vm.TurnOffRows.Select(r => r.FirstColumn).ToList(),


                    // ------------------ 映射参数表 ------------------
                    DetailedPerformance = vm.ParameterRows.Select(row => new ParameterDetail
                    {
                        Name = row.Name,
                        Unit = row.Unit,
                        Groups = row.Groups.Select(group => new PerformanceGroup
                        {
                            // 注意：这里使用 _rawValueMin/Max 来获取完整的、未被 UI 隐藏的值
                            Min = group.Min, // 假设您能访问到原始值
                            Typ = group.Type,
                            Max = group.Max  // 假设您能访问到原始值
                        }).ToList()
                    }).ToList()



                }
            };

            return productData;
        }

        /// <summary>
        /// 从 JSON 字符串中加载数据到 ViewModel。
        /// </summary>
        /// <param name="jsonString">包含 ProductData 结构的完整 JSON 字符串。</param>
        /// <param name="vm">要加载数据的目标 ViewModel 实例。</param>
        public void LoadProductDataToViewModel(string jsonString)
        {
            if (string.IsNullOrWhiteSpace(jsonString))
            {
                // 处理空字符串的情况
                return;
            }

            // 1. 反序列化 JSON 字符串为 C# 对象
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true // 允许属性名不区分大小写
            };

            ProductData productData;
            try
            {
                productData = JsonSerializer.Deserialize<ProductData>(jsonString, options);
            }
            catch (JsonException ex)
            {
                // 捕获和处理 JSON 解析错误
                Console.WriteLine($"JSON Deserialization Error: {ex.Message}");
                return;
            }

            if (productData?.Tables == null)
            {
                // 数据无效
                return;
            }

            // 2. 将反序列化后的数据映射回 ViewModel 集合
            MapTablesToViewModel(productData.Tables);
        }

        private void MapTablesToViewModel(ParamTables tables)
        {
            // A. 清空 ViewModel 集合（确保数据不会重复或混淆）
            vm.BasicParameterRows.Clear();
            vm.FeatureParameterRows.Clear();
            vm.ParameterRows.Clear();
            vm.AbsoluteRatingsRows.Clear();
            vm.CurrentVdVgTable.Clear();
            vm.NotesTable.Clear();
            vm.Description1Rows.Clear();
            vm.Description2Rows.Clear();
            vm.TurnOnRows.Clear(); // 注意：您的序列化方法中没有 TurnOnRows 的 Clear，但反序列化时应清空
            vm.TurnOffRows.Clear();


            // ------------------ 映射两列数据表 (Key/Value -> Name/Info) ------------------

            // 假设 MapParameterItemsToFeatureRows(List<ParameterItem> items, ObservableCollection<FeatureParameterRow> target) 已经实现
            MapParameterItemsToFeatureRows(tables.BasicParameters, vm.BasicParameterRows);
            MapParameterItemsToFeatureRows(tables.FeatureParameters, vm.FeatureParameterRows);
            MapParameterItemsToFeatureRows(tables.AbsoluteMaximumRatings, vm.AbsoluteRatingsRows);

            // ------------------ 映射 Description1 (Component/Description -> Name/Info) ------------------
            if (tables.Description1 != null)
            {
                foreach (var item in tables.Description1)
                {
                    // 假设 DescriptionItem1 的 Key 映射到 FeatureParameterRow 的 Name，Value 映射到 Info
                    vm.Description1Rows.Add(new FeatureParameterRow
                    {
                        Name = item.Component, // JSON Component 映射到 ViewModel Name
                        Info = item.Description // JSON Description 映射到 ViewModel Info
                    });
                }
            }

            // ------------------ 映射参数表 (DetailedPerformance) ------------------

            if (tables.DetailedPerformance != null)
            {
                foreach (var detail in tables.DetailedPerformance)
                {
                    var parameterRow = new ParameterRow
                    {
                        Name = detail.Name,
                        Unit = detail.Unit,
                    };

                    // 映射 Groups
                    if (detail.Groups != null)
                    {
                        foreach (var group in detail.Groups)
                        {
                            var minMaxGroup = new MinMaxTypeGroup();

                            // ? 关键：设置原始 Min/Max 值（假设 SetRawValues 存在）
                            minMaxGroup.SetRawValues(group.Min, group.Max);

                            // 设置 Type (Typ)
                            minMaxGroup.Type = group.Typ;

                            parameterRow.Groups.Add(minMaxGroup);
                        }
                    }
                    vm.ParameterRows.Add(parameterRow);
                }
            }

            // ------------------ 映射三列数据表 (SupplyCurrentVdVg) ------------------
            if (tables.SupplyCurrentVdVg != null)
            {
                foreach (var item in tables.SupplyCurrentVdVg)
                {
                    // 假设 TreeColumnTableRow 用于三列数据
                    vm.CurrentVdVgTable.Add(new TreeColumnTableRow
                    {
                        FirstColumn = item.VD,
                        SecondColumn = item.VG,
                        ThirdColumn = item.Current_mA // JSON 中 Current_mA 对应 ViewModel 第三列
                    });
                }
            }

            // ------------------ 映射三列数据表 (Description2) ------------------
            if (tables.Description2 != null)
            {
                foreach (var item in tables.Description2)
                {
                    vm.Description2Rows.Add(new TreeColumnTableRow
                    {
                        FirstColumn = item.Pin,
                        SecondColumn = item.Function,
                        ThirdColumn = item.Detail
                    });
                }
            }


            // ------------------ 映射单列数据表 (Notes, TurnOnProcedure, TurnOffProcedure) ------------------

            // Notes
            if (tables.Notes != null)
            {
                foreach (var note in tables.Notes)
                {
                    // 假设 OneColumnTableRow 用于单列数据
                    vm.NotesTable.Add(new OneColumnTableRow { FirstColumn = note });
                }
            }

            // TurnOnProcedure
            if (tables.TurnOnProcedure != null)
            {
                foreach (var item in tables.TurnOnProcedure)
                {
                    vm.TurnOnRows.Add(new OneColumnTableRow { FirstColumn = item });
                }
            }

            // TurnOffProcedure
            if (tables.TurnOffProcedure != null)
            {
                foreach (var item in tables.TurnOffProcedure)
                {
                    vm.TurnOffRows.Add(new OneColumnTableRow { FirstColumn = item });
                }
            }
        }

        /// <summary>
        /// 将 ParameterItem 列表反向映射到 FeatureParameterRow 的 ObservableCollection 中。
        /// </summary>
        /// <param name="items">从 JSON 反序列化得到的 ParameterItem 列表。</param>
        /// <param name="targetCollection">要加载数据的 ViewModel 目标集合。</param>
        private void MapParameterItemsToFeatureRows(
            List<ParameterItem> items,
            ObservableCollection<FeatureParameterRow> targetCollection)
        {
            // 1. 清空目标集合，准备加载新数据（重要步骤）
            targetCollection.Clear();

            // 2. 检查源数据是否有效
            if (items == null)
            {
                return;
            }

            // 3. 映射并添加到目标集合
            var newRows = items.Select(item => new FeatureParameterRow
            {
                Name = item.Key,    // JSON Key 映射到 ViewModel Name
                Info = item.Value   // JSON Value 映射到 ViewModel Info
            });

            foreach (var row in newRows)
            {
                targetCollection.Add(row);
            }
        }
    }

    public  class TablesModel : ObeservableObject
    {
        public ObservableCollection<ParameterRow> ParameterRows { get; set; } = new ObservableCollection<ParameterRow>();
        public ObservableCollection<FeatureParameterRow> FeatureParameterRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> BasicParameterRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<FeatureParameterRow> AbsoluteRatingsRows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<TreeColumnTableRow> CurrentVdVgTable { get; set; } = new ObservableCollection<TreeColumnTableRow>();
        public ObservableCollection<OneColumnTableRow> NotesTable { get; set; } = new ObservableCollection<OneColumnTableRow>();
        public ObservableCollection<FeatureParameterRow> Description1Rows { get; set; } = new ObservableCollection<FeatureParameterRow>();
        public ObservableCollection<TreeColumnTableRow> Description2Rows { get; set; } = new ObservableCollection<TreeColumnTableRow>();
        public ObservableCollection<OneColumnTableRow> TurnOnRows { get; set; } = new ObservableCollection<OneColumnTableRow>();
        public ObservableCollection<OneColumnTableRow> TurnOffRows { get; set; } = new ObservableCollection<OneColumnTableRow>();

    }


    public class MinMaxTypeGroup : INotifyPropertyChanged
    {
        private string _rawValueMin;
        private string _rawValueMax;

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
        public void UpdateDisplay(bool isMinVisible, bool isMaxVisible)
        {
            // 1. Min 属性的显示逻辑
            Min = isMinVisible ? _rawValueMin : "";
            //OnPropertyChanged(nameof(Min));

            // 2. Max 属性的显示逻辑
            Max = isMaxVisible ? _rawValueMax : "";
            //OnPropertyChanged(nameof(Max));
        }

        // ... 确保 SetRawValues 也调用 UpdateDisplay(true, true) ...
        public void SetRawValues(string min, string max)
        {
            _rawValueMin = min;
            _rawValueMax = max;
            UpdateDisplay(true, true); // 初始默认全部显示
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

    public class ParameterRow : ObeservableObject
    {
        public string Name { get; set; }
        public string Unit { get; set; }
        public ObservableCollection<MinMaxTypeGroup> Groups { get; set; } = new ObservableCollection<MinMaxTypeGroup>();

        private bool _isMinVisible = true;
        public bool IsMinVisible
        {
            get { return _isMinVisible; }
            set
            {
                Console.WriteLine($"[DEBUG] IsMinVisible SET to {value} on row: {Name}");
                if (_isMinVisible != value)
                {
                    _isMinVisible = value;
                    RaisePropertyChanged(nameof(IsMinVisible));

                    // 联动核心：通知所有 Groups 对象的 Min 属性更新显示
                    foreach (var group in Groups)
                    {
                        group.UpdateDisplay(value, IsMaxVisible);
                    }
                }
            }
        }

        // 【新增属性二：控制 Max 列的显示状态】
        private bool _isMaxVisible = true;
        public bool IsMaxVisible
        {
            get { return _isMaxVisible; }
            set
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] IsMinVisible SET to {value} on row: {Name}");
                if (_isMaxVisible != value)
                {
                    _isMaxVisible = value;
                    RaisePropertyChanged(nameof(IsMaxVisible));

                    // 联动核心：通知所有 Groups 对象的 Max 属性更新显示
                    foreach (var group in Groups)
                    {
                        group.UpdateDisplay(IsMinVisible, value);
                    }
                }
            }
        }

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



    public class TreeColumnTableRow : ObeservableObject
    { 
       public String FirstColumn { set; get; }
    
       public String SecondColumn { set; get; }

       public String ThirdColumn { set; get; }
    
    }

    public class OneColumnTableRow : ObeservableObject
    {
        public String FirstColumn { set; get; }

    }

}

           