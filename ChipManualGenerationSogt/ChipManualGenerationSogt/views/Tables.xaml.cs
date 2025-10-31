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
    /// Tables.xaml �Ľ����߼�
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
                    Console.Write($"{value[i, j],-15}"); // ÿ�п�ȹ̶��������
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
                Type = "��1.0"
            };
            var item31 = new MinMaxType()
            {
                Type = "��1.0"
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


            // ʾ�����������������
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
            FeatureAddParameterRow("Gain Flatness", "��2.5dB Typical");
            FeatureAddParameterRow("Noise Figure", "4.5dB Typical");
            FeatureAddParameterRow("P1dB", "12dBm Typical");
            FeatureAddParameterRow("Power Supply", "VD=+4V@119mA ,VG=-0.4V");
            FeatureAddParameterRow("Input/Output", "50��");
            FeatureAddParameterRow("Chip Size", "1.766 x 2.0 x 0.05mm ");


            
        }

        private void InitializeBasicParameterRows()
        {
            vm.BasicParameterRows.Clear();
            BasicAddParameterRow("Manual PN", "MML806");
            BasicAddParameterRow("Version", "V3.0.0");
            BasicAddParameterRow("Product Name", "��2.5dB Typical");
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

            // ��һ�У���������
            parametersDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Parameters",
                Binding = new Binding("Name"),
                IsReadOnly = true
            });

            // ��̬��� Min/Type/Max ��
            if (vm.ParameterRows.Count > 0)
            {
                int groupCount = vm.ParameterRows[0].Groups.Count; // ��������������һ��

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

            // ���һ�У���λ
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
            // ? ������Show Max CheckBox ��
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

            // ��һ�У���������
            featureDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Feature",
                Binding = new Binding("Name"),
                //IsReadOnly = true
            });


            // ���һ�У���λ
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

            // ��һ�У���������
            basicDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Feature",
                Binding = new Binding("Name"),
                //IsReadOnly = true
            });


            // ���һ�У���λ
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
        // ��ѡ�����·�����
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
        }// ͨ�ò��뷽��
        private void InsertRowAtOffset(int offset)
        {
            if (featureDataGrid.SelectedItems.Count == 0)
                return;

            // ��ȡ��һ��ѡ���֧�ֶ�ѡ����ֻȡ��һ����
            var selectedItem = featureDataGrid.SelectedItems[0];

            // �ҵ�����Դ�����е�����
            int currentIndex = vm.FeatureParameterRows.IndexOf(selectedItem as FeatureParameterRow);
            if (currentIndex == -1)
                return;

            // �������λ��
            int insertIndex = currentIndex + offset;

            // �߽紦������С�� 0
            if (insertIndex < 0)
                insertIndex = 0;

            // ��������
            vm.FeatureParameterRows.Insert(insertIndex, new FeatureParameterRow());

            // ��ѡ���Զ�ѡ���²������
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
            // ������������
            HideAllContent();

            // ����ѡ��ı��������ʾ��Ӧ����
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

        // ����������ö��

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

            // ��������ת��Ϊ TableType
            TableType tableType = GetTableTypeFromContent(content);

            // ��ʾ��Ӧ�ı��
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
            // �������б��
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


            // ��������ͼƬ����
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

            // ����ͼ��
            WpfPlot1.Visibility = Visibility.Collapsed;
        }

        private void InitializeAllTables()
        {
            // ���������⴦��
            InitializeParameterRows(_currentGroupCount);
            RebuildDataGridColumns();

            // �������ʹ��ͳһ�ĳ�ʼ������
            var basicTable = new List<(string, string)>
                            {
                                ("{Top PN}", "MML806"),
                                ("{Version}", "V1.0.0"),
                                ("{Product Name}", "GaAs MMIC Low Noise Amplifier"),
                                ("{Frequency Range}", "0-10GHz"),
                                ("{right bar info}", "GaAs Low Noise Amplifier MMIC 45 �C 90GHz")
                            }; 
            InitializeTableData(vm.BasicParameterRows, basicTable);
            var featureTable =new List<(string, string)>
    {
        ("Frequency", "45-90GHz"),
        ("Small Signal Gain", "15dB Typical"),
        ("Gain Flatness", "��2.5dB Typical"),
        ("Noise Figure", "4.5dB Typical"),
        ("P1dB", "12dBm Typical"),
        ("Power Supply", "VD=+4V@119mA ,VG=-0.4V"),
        ("Input/Output", "50��"),
        ("Chip Size", "1.766 x 2.0 x 0.05mm")
    };

            InitializeTableData(vm.FeatureParameterRows, featureTable);

            var rateTable = new List<(string, string)>
            {
                ("Drain Bias Voltage (VD)", "+4.5V"),
        ("Gate Bias Voltage (VG)", "-2V to 0V"),
        ("RF Input Power (RFIN)", "+15dBm"),
        ("Continuous Pdiss (T = 85 ��C)\n (derate 6.1mW/��C above 85 ��C) ", "175��C"),
        ("Thermal Resistance\n (channel to die bottom)", "0.55W"),
        ("Operating Temperature", "-55��C to +85 ��C"),
        ("Storage Temperature", "-65��C to +150 ��C"),
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
        ( "1. Die thickness: 50��m"),
        ( "2. VD bond pad is 75*75��m?"),
        ( "3. VG bond pad is 75*75��m?"),
        ( "4. RF IN/OUT bond pad is 50*86��m?"),
        ( "5. Bond pad metalization: Gold"),
        ( "6. Backside metalization: Gold"),
    };

            InitializeTable(vm.NotesTable, notesTable);

            var d1Table =new List<(string, string)>
    {
        ("C1", "100pF Example: Skyworks Part: SC10002430"),
        ("C2", "0.01��F Example: TDK Part:C1005X7R1H103K050BB (0402)"),
        ("C3", "0.1��F Example: TDK Part:C1005X7R1H104K050BB (0402)"),
        ("R1", "100�� Example: Yageo Part:SR0402FR-7T10RL"),
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
                        "2.    Set the gate bias voltages,VG to -2V.",
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
                        "4.    Increase the gate bias voltages to achieve a quiescent supply current of 82 mA.",
                        "5.    Increase the all gate bias voltages to 0 V.",
                    };
            InitializeTable(vm.TurnOffRows, t2Table);

            // Ϊ���б���ؽ��У����˲�����
            RebuildAllDataGridColumns();
        }

        // ͳһ�ı�����ݳ�ʼ������
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
        /// �������еı�����ݳ�ʼ��
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
        // �����������ݶ���
        public List<(string name, string info)> GetBasicTableData()
        {
            return GetTwoMetaTable(vm.BasicParameterRows);
        }
        
        public List<(string name, string info)> GetFeatureTableData()
        {
            return GetTwoMetaTable(vm.FeatureParameterRows);
        }

        /// <summary>
        /// ����������⣬ ����ֵҪ���⴦��
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
        /// ��ѹ�� ������
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
            // ������ʹ���������ؽ�
            RebuildDataGridColumns();

            // �������ʹ��ͳһ���ؽ�
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
            // ���� vm �� _currentGroupCount �Ѿ�������

            // 1. �ⲿ���������Լ��
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }
            if (data.Count != vm.ParameterRows.Count)
            {
                throw new Exception("The parameter table data rows count is not right.");
            }
            // ȷ������������һ�У��Ա�������
            if (data.Count > 0 && data[0].Count != 3 * _currentGroupCount)
            {
                throw new Exception("The parameter table data columns count is not right.");
            }

            // 2. ���������� (i)
            for (int i = 0; i < data.Count; i++)
            {
                // 3. ���������� (j)��ʹ�ñ�׼�� for ѭ�������� j += 3
                // j ����ǰ�����ʼ������ (Min, Type, Max)
                for (int j = 0; j < data[i].Count; j += 3)
                {
  
                    var targetGroup = vm.ParameterRows[i].Groups[j / 3];
                    double valMin = data[i][j];      // Min ֵ
                    double valType = data[i][j + 1]; // Type/Nominal ֵ
                    double valMax = data[i][j + 2];  // Max ֵ
                    if (i == 0)
                    {

                        targetGroup.Min = "";
                        targetGroup.Max = "";
                        targetGroup.Type = valMin.ToString("f0") + " - " + valMax.ToString("f0");
                    }
                    // i=2 �������߼���Type ����Ϊ ��ֵ
                    else if (i == 2)
                    {
                        
                        string tmp = data[i][j+1].ToString("f1");
                        targetGroup.Min = "";
                        targetGroup.Type = "��" + tmp; 
                        targetGroup.Max = "";
                    }
                    // ������ i!=0 �� i!=2 ��ͨ���߼�
                    else
                    {
                        
                        //double valMin = data[i][j];      // Min ֵ
                        //double valType = data[i][j + 1]; // Type/Nominal ֵ
                        //double valMax = data[i][j + 2];  // Max ֵ

                        
                        //targetGroup.Min = valMin == 0 ? "" : valMin.ToString("f1");

                        
                        targetGroup.Type = valType == 0 ? "" : valType.ToString("f1");

                        
                       // targetGroup.Max = valMax == 0 ? "" : valMax.ToString("f1");


                        targetGroup.SetRawValues(valMin.ToString("f1"), valMax.ToString("f1"));
                    }
                }
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
            // 1. Min ���Ե���ʾ�߼�
            Min = isMinVisible ? _rawValueMin : "";
            //OnPropertyChanged(nameof(Min));

            // 2. Max ���Ե���ʾ�߼�
            Max = isMaxVisible ? _rawValueMax : "";
            //OnPropertyChanged(nameof(Max));
        }

        // ... ȷ�� SetRawValues Ҳ���� UpdateDisplay(true, true) ...
        public void SetRawValues(string min, string max)
        {
            _rawValueMin = min;
            _rawValueMax = max;
            UpdateDisplay(true, true); // ��ʼĬ��ȫ����ʾ
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

                    // �������ģ�֪ͨ���� Groups ����� Min ���Ը�����ʾ
                    foreach (var group in Groups)
                    {
                        group.UpdateDisplay(value, IsMaxVisible);
                    }
                }
            }
        }

        // ���������Զ������� Max �е���ʾ״̬��
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

                    // �������ģ�֪ͨ���� Groups ����� Max ���Ը�����ʾ
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

           