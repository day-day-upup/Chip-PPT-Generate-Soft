using ChipManualGenerationSogt;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using OpenTK.Input;
using ScottPlot;
using ScottPlot.Finance;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
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
using static ChipManualGenerationSogt.Curves;

namespace ChipManualGenerationSogt
{
    /// <summary>
    /// Filter.xaml �Ľ����߼�
    /// </summary>
    public partial class Filter : UserControl
    {
        FilterModel vm ;
        public event EventHandler OnQueryDatabaseFinished;
        public AmpfilierFilesbyGroup FilesByGroup ;
        List<TestRecord> _records;

        public bool QueryFinished { get; set; } = false;
        public bool CalculateFinished { get; set; } = false;
        public Filter()
        {
            InitializeComponent();
            vm = new FilterModel();
            DataContext = vm;
            _records = new List<TestRecord>();
            vm.PN = "L004X";
            vm.SN = "L004X";
            //vm.ManualPptName = "L004X";

            vm.Levels.Add("L004X");
            vm.Levels.Add("L005X");
            vm.Levels.Add("L006X");
            combox.ItemsSource = new List<string> { "1", "2", "3" };
        }


        public FileterConditionModel GetFileterCondition()

        {
            var condition = new FileterConditionModel
            {
                PN = vm.PN,
                ON  =vm.SN,
                Min= Convert.ToDouble(vm.MinFrequency),
                Max = Convert.ToDouble(vm.MaxFrequency)
                

            };
            if (combox.SelectedIndex == 0)
            {
                condition.FreqBands.Add(vm.Band1);
            }
            else if (combox.SelectedIndex == 1)
            {
                condition.FreqBands.Add(vm.Band1);
                condition.FreqBands.Add(vm.Band2);
            }
            else if (combox.SelectedIndex == 2)
            {
                condition.FreqBands.Add(vm.Band1);
                condition.FreqBands.Add(vm.Band2);
                condition.FreqBands.Add(vm.Band3);
            }

                string[] levels = vm.SelectedEntry.TrimEnd(';').Split(';');
            foreach (var item in levels)
            {
                condition.VD_VG_Conditon.Add(item); 
            }
            return condition;

        }

        public void SetFileterCondition(FileterConditionModel condition)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                vm.PN = condition.PN;
                vm.SN = condition.ON;
                vm.MinFrequency = condition.Min.ToString();
                vm.MaxFrequency = condition.Max.ToString();
                vm.StartDateTime = condition.StartDateTime;
                vm.EndDateTime = condition.StopDateTime;
                vm.SelectedEntry = "";
                foreach (var item in condition.VD_VG_Conditon)
                {

                    vm.SelectedEntry += item + ";";
                }
                combox.SelectedIndex = condition.FreqBands.Count-1;
                for (int i = 0; i < condition.FreqBands.Count; i++)
                {
                    if (i == 0)
                        vm.Band1 = condition.FreqBands.ElementAt(i);
                    else if (i == 1)
                        vm.Band2 = condition.FreqBands.ElementAt(i);
                    else if (i == 2)
                        vm.Band3 = condition.FreqBands.ElementAt(i);
                }
            });
        
        }
        public FilterModel getViewMode()
        {
            return vm;
        }
        public async void Btn_Next_Clicked(object sender, RoutedEventArgs e)
        {
            try
            {
                System.IO.Directory.Delete(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiedReports"), true);
            }
            catch
            { 
            }
            // ���ÿ���ֶ��Ƿ��д���
            string snError = vm[nameof(vm.SN)];
            string pnError = vm[nameof(vm.PN)];
            //string pptError = vm[nameof(vm.ManualPptName)];
            string startError = vm[nameof(vm.StartDateTime)];
            string endError = vm[nameof(vm.EndDateTime)];
            //Console.WriteLine(vm.SN);

            // ������κ�һ������
            if (!string.IsNullOrEmpty(snError))
            {
                MessageBox.Show(snError, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                snTextBox.Focus();
                return;
            }

            if (!string.IsNullOrEmpty(pnError))
            {
                MessageBox.Show(pnError, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                pnTextBox.Focus();
                return;
            }

            //if (!string.IsNullOrEmpty(pptError))
            //{
            //    MessageBox.Show(pptError, "��֤����", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    pptNameTextBox.Focus();
            //    return;
            //}

            if (!string.IsNullOrEmpty(startError))
            {
                //MessageBox.Show(startError, "��֤����", MessageBoxButton.OK, MessageBoxImage.Warning);
                //StartPicker.Focus();
                return;
            }

            if (!string.IsNullOrEmpty(endError))
            {
                //MessageBox.Show(endError, "��֤����", MessageBoxButton.OK, MessageBoxImage.Warning);
                //EndPicker.Focus();
                return;
            }

            // ? �����ֶζ�ͨ����֤������ִ������߼�
            //MessageBox.Show("�����ֶ�����д���������ɱ��棡", "�ɹ�", MessageBoxButton.OK, MessageBoxImage.Information);
            vm.IsBusy =true;
            vm.BusyMessage = "On executing query database...";
            try
            {
                var sqlServer = new TaskRepository();
                string message = $"prepare to prodece data...: conditon: StartDateTime:{vm.StartDateTime}, EndDateTime:{vm.EndDateTime}, PN:{vm.PN}, SN:{vm.SN}";
                var logmodel = new LogModel
                {
                    UserName = Global.User.UserName,
                    Message = message,
                    TimeStamp = DateTime.Now,
                    Level = LogLevels.Info
                };

                await sqlServer.InsertLogAsync(logmodel);
                await Task.Run(() => QueryDatabase(vm.PN, vm.SN, vm.StartDateTime, vm.EndDateTime));

                await Task.Run(() => FileSCopyGourp()); 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                try
                {
                    // ɾ��Ŀ¼
                    //System.IO.Directory.Delete(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiedReports"), true);
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        btn_excute.Content = "Re Process";
                        //OnQueryDatabaseFinished?.Invoke(this, EventArgs.Empty);
                        secondPartStackPanel.Visibility = Visibility.Visible;
                        //AddLevels(legends.ToList());

                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                vm.IsBusy = false;
                QueryFinished= true;
            }

        }

        private async void QueryDatabase(string pn, string on, DateTime? startdatetime, DateTime? stopdatetime)
        {
            #region ���ݿ��ѯ
            string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";

            var repo = new TestRecordRepository(connStr);
            //var records = repo.GetRecordsByPN(
            //    pnList: new[] { "L004X"},
            //    startTime: new DateTime(2025, 10, 1),
            //    endTime: DateTime.Now
            //);
            string keyword = pn;
            _records = repo.GetRecordsByPN(
                keywords: new[] { keyword }
            );
            if (_records.Count == 0)
            {
                var sqlServer = new TaskRepository();
                string message = "Not Found Any Record From Database";
                var logmodel = new LogModel
                {
                    UserName = Global.User.UserName,
                    Message = message,
                    TimeStamp = DateTime.Now,
                    Level = LogLevels.Info
                };
                await sqlServer.InsertLogAsync(logmodel);
                MessageBox.Show("Not Found Any Record From Database��", "Tips", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
                Console.WriteLine($"Found {_records.Count} records.");
            foreach (var r in _records)
            {
                Console.WriteLine($"{r.ID} | {r.PN} | {r.TestTime}");
            }

            var sqlServer1 = new TaskRepository();
            string message1 = "Access Database Success";
            var logmodel1 = new LogModel
            {
                UserName = Global.User.UserName,
                Message = message1,
                TimeStamp = DateTime.Now,
                Level = LogLevels.Info
            };
            await sqlServer1.InsertLogAsync(logmodel1);
            #endregion

            Application.Current.Dispatcher.Invoke(() =>
            {
                //btn_excute.Content = "Re Process";
                ////OnQueryDatabaseFinished?.Invoke(this, EventArgs.Empty);
                //secondPartStackPanel.Visibility = Visibility.Visible;
                //AddLevels(legends.ToList());

            });
            
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //var selectedItem = (sender as ListBox).SelectedItem; // ����ѡ���ʱ���õķ�ʽ
            //if (selectedItem == null)
            //{
            //    return;
            //}
            //Console.WriteLine(selectedItem.ToString());
            //vm.SelectedEntry = selectedItem.ToString();


            ListBox listBox = sender as ListBox;
            if (listBox == null)
            {
                return;
            }

            // *** ��ȷ��������ʹ�� SelectedItems ***
            System.Collections.IList selectedItems = listBox.SelectedItems;

            // 1. ��ӡ����ѡ�е���
            Console.WriteLine("--- ��ǰѡ�е������� ---");
            foreach (var item in selectedItems)
            {
                Console.WriteLine(item.ToString());
            }
            Console.WriteLine("--------------------------");

            // 2. ���� ViewModel (vm)
            // ע�⣺�������ѡ�ж����� vm.SelectedEntry ����Ӧ����һ������/�б������ǵ����ַ�����

            // ʾ����������ѡ������ַ���ƴ������
            string allSelected = string.Join(", ", selectedItems.Cast<object>().Select(i => i.ToString()));
            // ������� ViewModel ��һ�����Կ��Խ�������ѡ�е���
            // vm.SelectedEntries = allSelected; 

            // �������ֻ�洢һ����ͨ���ǵ�һ����
            if (selectedItems.Count > 0)
            {
                vm.SelectedEntry = "";

                foreach (var item in selectedItems)
                {
                    vm.SelectedEntry += item.ToString() + ";";
                }
            }
            else
            {
                vm.SelectedEntry = null; // ���� string.Empty
            }
        }

        public void AddLevels(List<string> levels)
        {
            vm.Levels.Clear();
            foreach (var level in levels)
            { 
                vm.Levels.Add(level);
            
            }
        }

        public async void FileSCopyGourp( )
        {

            #region ��datapc03 �и����ļ�������
            string filePath = _records.ElementAt(0).PN;

            var copier = new NetworkFolderCopier();

            // ��ѡ���Զ�����־���������д���ļ���
            // copier.Log = msg => File.AppendAllText("copy.log", $"{DateTime.Now:HH:mm:ss} {msg}\n");

            try
            {
                //����ON�����·��
                copier.CopyMatchingSubFolders(
                    networkRoot: @"\\DATAPC03\RFAutoTestReport$\Chip Verification",
                    PN: vm.PN,
                    ON: vm.SN,
                    localTargetBase: System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiedReports"),
                    username: "",   // ���ձ�ʾʹ�õ�ǰ�û�ƾ��
                    password: ""
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"?? �����쳣: {ex.Message}");
            }
            #endregion
            var sqlServer1 = new TaskRepository();
            string message1 = "Access Storage Success";
            var logmodel1 = new LogModel
            {
                UserName = Global.User.UserName,
                Message = message1,
                TimeStamp = DateTime.Now,
                Level = LogLevels.Info
            };
            await sqlServer1.InsertLogAsync(logmodel1);

            #region �����ƹ������ļ����з���
            var finder = new TextFileFinder(
            rootDirectory: @"CopiedReports",
             extensions: new[] { ".txt", "s2p" }
            );
            //var finder = new TextFileFinder(); // ������ļ�������
            var allFiles = finder.FindAllTextFiles(); // �������·���б�

            // ��������ļ���������
            FilesByGroup = AmplifierFileProcessor.ProcessFiles(allFiles);
            var legends = new Collection<string>();
            string temperature = "";
            string elecParam = "";
            if (FilesByGroup.DataSparabyTemp.TryGetValue("25.0deg", out var s2pAt25))
            {
                foreach (var item in s2pAt25)
                {

                    CureveGenerateLengdText(item, out temperature, out elecParam);

                    legends.Add(elecParam);
                }
                //s2pAt25.ForEach(Console.WriteLine);
               
            }

            #endregion
            Application.Current.Dispatcher.Invoke(() =>
            {
                
                AddLevels(legends.ToList());

            });
        }

        public async  void Btn_Calcute_Click(object sender, RoutedEventArgs e)
        {
            while (!QueryFinished)
            {
                await Task.Delay(1000);
            }
            string message1 = "";
            
            if (vm.SelectedEntry == null || vm.SelectedEntry.Length == 0)
            {
                MessageBox.Show(" Please Select Parameter Level!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (combox.SelectedIndex == 0)
                {
                    if (vm.Band1 == null || vm.Band1.Length == 0)
                    {
                        MessageBox.Show(" Please Enter Frequency Band1 Range!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
                else if (combox.SelectedIndex == 1)
                {

                    if (vm.Band1 == null || vm.Band1.Length == 0)
                    {
                        MessageBox.Show(" Please Enter Frequency Band1 Range!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    if (vm.Band2 == null || vm.Band2.Length == 0)
                    {
                        MessageBox.Show(" Please Enter Frequency Band2 Range!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                }
                else if (combox.SelectedIndex == 2)
                {
                    if (vm.Band1 == null || vm.Band1.Length == 0)
                    {
                        MessageBox.Show(" Please Enter Frequency Band1 Range!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    if (vm.Band2 == null || vm.Band2.Length == 0)
                    {
                        MessageBox.Show(" Please Enter Frequency Band2 Range!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    if (vm.Band3 == null || vm.Band3.Length == 0)
                    {
                        MessageBox.Show(" Please Enter Frequency Band3 Range!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }


                message1 = $"condition:StartDateTime:SelectedEntry:{vm.SelectedEntry}, MinFrequency:{vm.MinFrequency}," +
                    $" MaxFrequency :{vm.MaxFrequency}, Bands:{combox.SelectedItem.ToString()}, Band1:{vm.Band1},Band2:{vm.Band2}," +
                    $" Band3:{vm.Band3}.Calculate success";
            });
            var sqlServer1 = new TaskRepository();
            //string message1 = $"condition:StartDateTime:SelectedEntry:{vm.SelectedEntry}, MinFrequency:{vm.MinFrequency}," +
            //    $" MaxFrequency :{vm.MaxFrequency}, Bands:{combox.SelectedItem.ToString()}, Band1:{ vm.Band1},Band2:{vm.Band2}," +
            //    $" Band3:{vm.Band3}.Calculate success";
            
            var logmodel1 = new LogModel
            {
                UserName = Global.User.UserName,
                Message = message1,
                TimeStamp = DateTime.Now,
                Level = LogLevels.Info
            };
            await sqlServer1.InsertLogAsync(logmodel1);
            QueryFinished = false;
            OnQueryDatabaseFinished?.Invoke(this, EventArgs.Empty);
        }

        private void CureveGenerateLengdText(string filePath, out string temperature, out string elecParam)
        {
            string baseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string[] parts = baseName.Split('_');

            // ��ȡ�¶ȣ��� "deg" ��β�ĶΣ�
            temperature = parts.FirstOrDefault(p => p.EndsWith("deg", StringComparison.OrdinalIgnoreCase))
                              ?? "UnknownTemp";

            // ��ȡ�������������� "VD=" �ĶΣ�
            elecParam = parts.FirstOrDefault(p => p.Contains("VD="))
                            ?? "UnknownParam";
        }

        private void combox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (combox.SelectedIndex == 0)
            {
                band1.Visibility = Visibility.Collapsed;
                band2.Visibility = Visibility.Collapsed;
            }
            else if (combox.SelectedIndex == 1)
            {
                band1.Visibility = Visibility.Visible;
                band2.Visibility = Visibility.Collapsed;
            }
            else if (combox.SelectedIndex == 2)
            {

                band1.Visibility = Visibility.Visible;
                band2.Visibility = Visibility.Visible;
            }

        }
    }



    public class FilterModel : ObeservableObject, IDataErrorInfo
    {
        // ========== �ַ����ֶ� ==========
        private string _pn;
        public string PN
        {
            get {return _pn; }
            set { this._pn = value; RaisePropertyChanged(nameof(PN)); }
        }

        private string _sn;
        public string SN
        {
            get { return _sn; }
            set { this._sn = value; RaisePropertyChanged(nameof(SN)); }
        }
        private string _standardLevel;

        public string SelectedEntry
        {
            get { return _standardLevel; }
            set { this._standardLevel = value; RaisePropertyChanged(); }
        }

        public ObservableCollection<string> Levels { set; get; } = new ObservableCollection<string>();

        // ========== ����ʱ���ֶΣ���Ϊ�գ�==========
        private DateTime? _startDateTime;
        public DateTime? StartDateTime
        {
            get { return _startDateTime; }
            set { this._startDateTime = value; RaisePropertyChanged(nameof(StartDateTime)); }
        }

        private DateTime? _endDateTime;
        public DateTime? EndDateTime
        {
            get { return _endDateTime; }
            set { this._endDateTime = value; RaisePropertyChanged(nameof(EndDateTime)); }
        }



        private string _maxFrequency="10.0";

        public string MaxFrequency
        {
            get { return _maxFrequency; }
            set { this._maxFrequency = value; RaisePropertyChanged(); }
        }

        private string _minFrequency ="0.0";

        public string MinFrequency
        {
            get { return _minFrequency; }
            set { this._minFrequency = value; RaisePropertyChanged(); }
        }

        // ========== IDataErrorInfo ��֤ ==========
        public string this[string columnName]
        {
            get
            {
                switch (columnName)
                {
                    case "PN":
                        return string.IsNullOrWhiteSpace(PN) ? "PN Cannot Be Empty" : null;

                    case "SN":
                        return string.IsNullOrWhiteSpace(SN) ? "SN Cannot Be Empty" : null;

                    //case "ManualPptName":
                    //    return string.IsNullOrWhiteSpace(ManualPptName) ? "PPT Name Cannot Be Empty" : null;

                    

                    default:
                        return null;
                }
            }
        }

        public string Error => null;


        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set { this._isBusy = value; RaisePropertyChanged(nameof(IsBusy)); }
        }


        private string busyMessage;
        public string BusyMessage
        {
            get { return busyMessage; }
            set { this.busyMessage = value; RaisePropertyChanged(nameof(BusyMessage)); }
        }

        private string _band1;
        public string Band1
        {
            get { return _band1; }
            set { this._band1 = value; RaisePropertyChanged(nameof(Band1)); }
        }

        private string _band2;
        public string Band2
        {
            get { return _band2; }
            set { this._band2 = value; RaisePropertyChanged(nameof(Band2)); }
        }


        private string _band3;
        public string Band3
        {
            get { return _band3; }
            set { this._band3 = value; RaisePropertyChanged(nameof(Band2)); }
        }
    }




}
