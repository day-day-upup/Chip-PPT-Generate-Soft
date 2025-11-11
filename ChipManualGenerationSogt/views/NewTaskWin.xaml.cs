using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Wpf.Toolkit;
using System.IO;
namespace ChipManualGenerationSogt
{
    /// <summary>
    /// NewTaskWin.xaml 的交互逻辑
    /// </summary>
    public partial class NewTaskWin : System.Windows.Controls.UserControl
    {
        //public event EventHandler AddTaskEvent;
        public event EventHandler BackEvent;
        const string ADD_TASK_STRING = "Add New Task";
        const string EDIT_TASK_STRING = "Update Task";
        TaskTableItem _currentTask;
        NewTaskWinModel vm;
        public NewTaskWin()
        {
            InitializeComponent();
            vm = new NewTaskWinModel();
            DataContext = vm;
            Init();
        }
        //public void Init()
        //{

        //    List<string> list = new List<string>();
        //    list.Add("MML806");
        //    var parent = new DeviceTreeViewItem
        //    {
        //        Content = "Amplifier",
        //        Visible = true,
        //    };
        //    var children = new ObservableCollection<DeviceTreeViewItem>(
        //        //list.Select(s => new DeviceTreeViewItem { Content = s, Parent = parent })
        //        list.Select(s => new DeviceTreeViewItem { Content = s, Parent = parent })
        //    );
        //    // 3. 将子节点集合赋给父节点
        //    parent.Children = children;

        //    // 4. 添加到根集合
        //    vm.DeviceTreeResources.Add(parent);

        //    vm.ParentObjectName = "Task Manager";

        //}
        public void ShowCurrentTaskConfigure(TaskTableItem task)
        {
            _currentTask = task;
            secondPartStackPanel.Visibility = Visibility.Visible;
            vm.TaskName = task.TaskName;
            vm.SelectedStatus = task.Status;
            vm.SelectedLevel = task.Level;
            vm.SelectedMajor = task.Major;
            vm.SelectedMinor = task.Minor;
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
            };
            var conditons = JsonSerializer.Deserialize<TaskFrequencyConfig>(task.Conditions, options);
            vm.SelectedEntry = conditons.SelectedEntry;
            vm.MinFrequency = conditons.MinFrequency;
            vm.MaxFrequency = conditons.MaxFrequency;
            
            
            if (conditons.Bands == "1")
            {
                combox.SelectedIndex = 0;
            }
            else if (conditons.Bands == "2")
            {
                combox.SelectedIndex = 1;
            }
            else if (conditons.Bands == "3")
            {
                combox.SelectedIndex = 2;
            }
           
            vm.ParameterITem = new ObservableCollection<object>(conditons.ParameterItems.Cast<object>());
            vm.ParametersSource.Clear();
            foreach (var item in conditons.ParameterItems)
            {
                vm.ParametersSource.Add(item);


            }
            string[] band1 = conditons.Band1.Split('-');
            string[] band2 = conditons.Band2.Split('-');
            string[] band3 = conditons.Band3.Split('-');
            vm.Band1MaxValue =Convert.ToDouble( band1[1].Trim());
            vm.Band1MinValue = Convert.ToDouble(band1[0].Trim());
            vm.Band2MaxValue = Convert.ToDouble(band2[1].Trim());
            vm.Band2MinValue = Convert.ToDouble(band2[0].Trim());
            vm.Band3MaxValue = Convert.ToDouble(band3[1].Trim());
            vm.Band3MinValue = Convert.ToDouble(band3[0].Trim());
            //vm.SelectedPPTModel = task.PPTModel;
            _button.Content = EDIT_TASK_STRING;


        }
        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {

        }

        private void Menu_Preview_Clicked(object sender, RoutedEventArgs e)
        {

        }
        private void CheckComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is CheckComboBox checkComboBox)
            {

                // 1. 查找 CheckComboBox 模板内部的 Popup 控件
                Popup popup = FindVisualChild<Popup>(checkComboBox);

                if (popup != null)
                {
                    // 2. 移除旧的订阅（防止重复）
                    popup.Opened -= Popup_Opened;

                    // 3. ? 订阅 Popup 的 Opened 事件
                    popup.Opened += Popup_Opened;

                    // (可选：如果你也需要收起事件)
                    // popup.Closed -= Popup_Closed;
                    // popup.Closed += Popup_Closed;
                }
            }
        }
        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child != null && child is T typedChild)
                {
                    // 如果找到目标类型，直接返回
                    return typedChild;
                }
                else
                {
                    // 递归查找子控件的子控件
                    T result = FindVisualChild<T>(child);
                    if (result != null)
                        return result;
                }
            }
            return null;
        }

        private void Popup_Opened(object sender, EventArgs e)
        {
            // 这里的代码会在 CheckComboBox 的下拉列表展开时执行
            Console.WriteLine("CheckComboBox 下拉列表已展开！执行动态加载或更新逻辑...");

            if (sender is Popup popup)
            {
                // 尝试向上查找 CheckComboBox 控件
                // 在 CheckComboBox 的模板中，Popup 可能是 CheckComboBox 视觉树的直接子元素，
                // 但更安全的方法是使用辅助函数向上查找特定类型。

                // 1. 获取 Popup 的 PlacementTarget，这通常是触发下拉的那个元素，但在这里可能不直接是 CheckComboBox。
                // 2. 向上遍历可视化树（更可靠）

                CheckComboBox parentCheckComboBox = FindParent<CheckComboBox>(popup);

                if (parentCheckComboBox != null)
                {
                    string comboBoxName = parentCheckComboBox.Name;

                    if (comboBoxName == "_cbMajor")
                    {
                        Console.WriteLine(" major");
                    }
                    else if (comboBoxName == "_cbMinor")
                    {
                        Console.WriteLine(" minor");
                    }
                    //    Console.WriteLine($"CheckComboBox 名称是: {comboBoxName}");
                    //Console.WriteLine("CheckComboBox 下拉列表已展开 (通过拦截内部 Popup 事件)。");

                    // ... 其他逻辑
                }
            }
        }

        private static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            // 从子元素开始，循环直到找到匹配的父元素或到达可视化树的根
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            while (parentObject != null)
            {
                if (parentObject is T parent)
                {
                    return parent;
                }

                // 继续向上查找
                parentObject = VisualTreeHelper.GetParent(parentObject);
            }
            return null;
        }

        private void Btn_SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
               
                dialog.Description = "Please Select A Folder";

                // （可选）设置初始选择的文件夹路径
                // dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

               
                DialogResult result = dialog.ShowDialog();

                // 检查用户是否点击了“确定”按钮
                if (result == DialogResult.OK)
                {
                    string selectedPath = dialog.SelectedPath;
                    //MessageBox.Show($"您选择的文件夹是: {selectedPath}");
                    //return selectedPath;
                    //secondPartStackPanel.Visibility = Visibility.Visible;
                    //vm.FolderPath = selectedPath;

                    SelectForder(selectedPath);
                }
                else
                {
                    //MessageBox.Show("您取消了文件夹选择。");
                    //return null;
                }
            }
        }

        private void CB_Bands_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = sender as System.Windows.Controls.ComboBox;
            if (comboBox == null)
            {
                return;
            }
            if (sender == null)
            {
                return;
            }
            if (comboBox.SelectedItem == null)
            {
                // 如果没有选中项，可以不执行后续逻辑，或者执行隐藏所有控件的操作

                return;
            }
            else
            {
                string selectedBand = (sender as System.Windows.Controls.ComboBox).SelectedItem.ToString();

                if (selectedBand == "1")
                {
                    BandsHidenAll();
             
                }

                else if (selectedBand == "2")
                {
                    BandsHidenAll();
                    if (_band2 != null)
                        _band2.Visibility = Visibility.Visible;
                }
                else if (selectedBand == "3")
                {
                    BandsHidenAll();
                    if (_band2 != null)
                        _band2.Visibility = Visibility.Visible;
                    if (_band3 != null)
                        _band3.Visibility = Visibility.Visible;

                }
            }
        
        }

        private async void Btn_Add_Click(object sender, RoutedEventArgs e)
        {
            string message;
            bool success = false;
            if (string.IsNullOrEmpty(vm.SelectedPPTModel) &&
                string.IsNullOrEmpty(vm.TaskName) &&
                string.IsNullOrEmpty(vm.PptName) &&
                string.IsNullOrEmpty(vm.SelectedStatus) &&
                string.IsNullOrEmpty(vm.SelectedLevel) &&
                string.IsNullOrEmpty(vm.SelectedMajor) &&
                string.IsNullOrEmpty(vm.SelectedMinor))
            {
                System.Windows.MessageBox.Show("Please fill in all the required fields!");
               
            }
            // 1?? 添加任务
            if (await AddTaskToDB())
            {
                // 2?? 上传文件夹
                string ftpRootPath = $"{vm.SelectedMajor}\\{(vm.SelectedMinor ?? "")}\\{vm.TaskName}";
                bool uploadResult = await FtpClient.UploadFolderAsync(vm.FolderPath, ftpRootPath);

                if (uploadResult)
                {
                    message = "Add Task Success!";
                    success = true;
                }
                else
                {
                    message = "Add Task Success, but Upload Folder Failed!";
                }
            }
            else
            {
                message = "Add Task Failed!";
            }

           
            System.Windows.MessageBox.Show(message, "Information", MessageBoxButton.OK, MessageBoxImage.Information);

          
            if (success)
                BackEvent?.Invoke(this, EventArgs.Empty);
        }


        private void Init()
        {
            NewTaskWinModel viewModel = vm;
            #region  从 JSON 文件加载初始数据
            // 假设 JSON 文件名为 initial_task_config.json 并且位于应用程序运行目录
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string JsonFileName = System.IO.Path.Combine(basePath, "resources", "files", "NewTask.Json");
            //const string JsonFileName = "NewTask.Json";

            if (!File.Exists(JsonFileName))
            {
                // 确保文件存在，否则进行错误处理或使用默认值
                Console.WriteLine($"Error: JSON file not found at {JsonFileName}");
                return;
            }

            try
            {
                string jsonString = File.ReadAllText(JsonFileName);
                // 使用 System.Text.Json 反序列化
                var root = JsonSerializer.Deserialize<RootObject>(jsonString);

                if (root?.InitialData != null)
                {
                    var sources = root.InitialData.Sources;
                    var defaults = root.InitialData.DefaultValues;

                    // --- A. 赋值数据源 (Sources) ---

                    // 为了简化代码，我们使用一个辅助方法来加载列表到 ObservableCollection
                    LoadSource(viewModel.StatusSource, sources.StatusSource);
                    LoadSource(viewModel.LevelSource, sources.LevelSource);
                    LoadSource(viewModel.MajorSource, sources.MajorSource);
                    LoadSource(viewModel.MinorSource, sources.MinorSource);
                    LoadSource(viewModel.PPTModelSource, sources.PPTModelSource);
                    //LoadSource(viewModel.ParametersSource, sources.ParametersSource);
                    LoadSource(viewModel.BandsSource, sources.BandsSource);

                    // --- B. 赋值默认值 (DefaultValues) ---

                    //viewModel.TaskName = defaults.TaskName;
                    //viewModel.FolderPath = defaults.FolderPath;
                    //viewModel.SelectedStatus = defaults.SelectedStatus;
                    //viewModel.SelectedLevel = defaults.SelectedLevel;
                    //viewModel.SelectedMajor = defaults.SelectedMajor;
                    //viewModel.SelectedMinor = defaults.SelectedMinor;
                    //viewModel.SelectedPPTModel = defaults.SelectedPPTModel;
                    //viewModel.SelectedEntry = defaults.SelectedEntry;

                    // Min/Max Frequency
                    //viewModel.MinFrequency = defaults.MinFrequency;
                    //viewModel.MaxFrequency = defaults.MaxFrequency;
                    

                    // Frequency Bands
                    //viewModel.Band1 = defaults.Band1;
                    //viewModel.Band2 = defaults.Band2;
                    //viewModel.Band3 = defaults.Band3;

                    // 【可选】设置 ParentObjectName 的默认值，但它通常来自 TreeView 的选择
                    // viewModel.ParentObjectName = "Default Parent"; 
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"JSON Deserialization Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred during initialization: {ex.Message}");
            }
            #endregion
            viewModel.MinFrequency = 0.0;
            viewModel.MaxFrequency = 10.0;
            vm.ParentObjectName = "New Task";

        }

        // 辅助方法：安全地将 List<string> 加载到 ObservableCollection<string>
        private void LoadSource(ObservableCollection<string> destination, List<string> source)
        {
            if (source != null)
            {
                destination.Clear();
                foreach (var item in source)
                {
                    destination.Add(item);
                }
            }
        }

        private void Border_DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            // 检查拖拽的数据对象中是否包含文件路径信息 (DataFormats.FileDrop)
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                // 如果是文件，允许拖放操作（显示为复制图标）
                e.Effects = System.Windows.DragDropEffects.Copy;
            }
            else
            {
                // 否则，不允许拖放操作（显示为禁止图标）
                e.Effects = System.Windows.DragDropEffects.None;
            }

            // 标记事件已处理，防止事件继续路由
            e.Handled = true;
        }

        private void Border_Drop(object sender, System.Windows.DragEventArgs e)
        {
            // 再次确认数据中包含文件路径
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                // 从 Data 对象中获取文件路径数组 (返回 string[])
                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);

                if (files != null && files.Length > 0)
                {
                    // 假设我们只关心第一个文件（如果你只希望用户拖拽一个文件）
                    string filePath = files[0];

                    //System.Windows.MessageBox.Show($"成功拖拽文件，路径是: {filePath}", "文件拖拽成功");
                    SelectForder(filePath);

                }
            }
        }

        private void SelectForder(string filePath)
        {
            secondPartStackPanel.Visibility = Visibility.Visible;
            vm.FolderPath = filePath;

            #region 将文件进行分类
            var finder = new TextFileFinder(
            rootDirectory: vm.FolderPath,
             extensions: new[] { ".txt", "s2p" }
            );
            //var finder = new TextFileFinder(); // 或你的文件查找器
            var allFiles = finder.FindAllTextFiles(); // 返回相对路径列表

            // 将里面的文件处理并分组
           var  FilesByGroup = AmplifierFileProcessor.ProcessFiles(allFiles);
            //var legends = new Collection<string>();
            string temperature = "";
            string elecParam = "";
            //将25deg的提取出来
            if (FilesByGroup.DataSparabyTemp.TryGetValue("25.0deg", out var s2pAt25))
            {
                foreach (var item in s2pAt25)
                {

                    CureveGenerateLengdText(item, out temperature, out elecParam);
                    
                    if (elecParam.StartsWith("-"))
                    {
                        // 如果以 '-' 开头，则从索引 1 开始截取字符串（即删除第一个字符）
                        elecParam = elecParam.Substring(1);
                    }
                   
                        vm.ParametersSource.Add(elecParam.Replace('&', ','));
                }
                //s2pAt25.ForEach(Console.WriteLine);

            }

            combox.SelectedIndex = 1;
            #endregion


        }

        private void CureveGenerateLengdText(string filePath, out string temperature, out string elecParam)
        {
            string baseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string[] parts = baseName.Split('_');

            // 提取温度（以 "deg" 结尾的段）
            temperature = parts.FirstOrDefault(p => p.EndsWith("deg", StringComparison.OrdinalIgnoreCase))
                              ?? "UnknownTemp";

            // 提取电气参数（包含 "VD=" 的段）
            elecParam = parts.FirstOrDefault(p => p.Contains("VD="))
                            ?? "UnknownParam";
        }
        /// <summary>
        /// 控制说有的频段显示条目隐藏起来
        /// </summary>
        private void BandsHidenAll()
        {
           if(_band2!=null)
            _band2.Visibility = Visibility.Collapsed;
           if(_band3!=null)
            _band3.Visibility = Visibility.Collapsed;

        }
        private void Btn_Preview_Click(object sender, RoutedEventArgs e)
        {
            BackEvent?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// 添加任务到数据库
        /// </summary>
        private async Task<bool> AddTaskToDB()
        {
            bool success = false;
            string content = _button.Content.ToString();
            if (content == EDIT_TASK_STRING)
            {
                var taskHelper = new TaskSqlServerRepository();
                var allTasks = await taskHelper.GetAllTasksAsync();
                int id = allTasks.Count() + 1;
                var taskItem = await taskHelper.GetTaskByIdAsync(_currentTask.ID);
 
                vm.Band1 = vm.Band1MinValue.ToString() + " - " + vm.Band1MaxValue.ToString();
                vm.Band2 = vm.Band2MinValue.ToString() + " - " + vm.Band2MaxValue.ToString();
                vm.Band3 = vm.Band3MinValue.ToString() + " - " + vm.Band3MaxValue.ToString();
                var dataToSerialize = new TaskFrequencyConfig
                {
                    //ParameterItems = vm.ParameterITem.Cast<object>().ToList(),
                    ParameterItems = vm.ParameterITem.Cast<string>().ToList(),
                    SelectedEntry = vm.SelectedEntry,
                    MaxFrequency = vm.MaxFrequency,
                    MinFrequency = vm.MinFrequency,
                    Bands = vm.Bands,
                    Band1 = vm.Band1,
                    Band2 = vm.Band2,
                    Band3 = vm.Band3
                };
                taskItem.Conditions = JsonSerializer.Serialize(dataToSerialize);
                taskItem.TaskName = vm.TaskName;

                taskItem.Status = vm.SelectedStatus;
                taskItem.Level = vm.SelectedLevel;
                taskItem.Major = vm.SelectedMajor;
                taskItem.Minor = vm.SelectedMinor;
                var result = await taskHelper.UpdateTaskAsync(taskItem);

                #region log record
                var logRecord = new LogModel
                {
                    TimeStamp = DateTime.Now,
                    UserName = Global.User.UserName,
                    TaskId = id.ToString(),
                    TaskName = vm.TaskName,
                    PN = null,
                    Level = vm.SelectedLevel,
                    SN = null,

                    ChipNumber = vm.SelectedPPTModel
                };
                logRecord.Message += "Update Task Success!";

           
                logRecord.Message += JsonSerializer.Serialize(taskItem);

                await LogRepository.InsertLogAsync(logRecord);

             
                #endregion
            }
            else if (content == ADD_TASK_STRING)
            {
                try
                {
                    var taskHelper = new TaskSqlServerRepository();
                    var allTasks = await taskHelper.GetAllTasksAsync();
                    int id = allTasks.Count() + 1;
                    var taskItem = new TaskSqlServerModel
                    {
                        ID = id,  //自增 不用显示使用
                        PPTModel = vm.SelectedPPTModel,
                        TaskName = vm.TaskName,
                        PptName = vm.PptName,
                        Status = vm.SelectedStatus,
                        Level = vm.SelectedLevel,
                        Major = vm.SelectedMajor,
                        Minor = vm.SelectedMinor,
                        StartDate = DateTime.Now,
                        EndDate = null,
                        DataStatus = false,
                        FilesStatus = false
                    };
                    vm.Band1 = vm.Band1MinValue.ToString() + " - " + vm.Band1MaxValue.ToString();
                    vm.Band2 = vm.Band2MinValue.ToString() + " - " + vm.Band2MaxValue.ToString();
                    vm.Band3 = vm.Band3MinValue.ToString() + " - " + vm.Band3MaxValue.ToString();
                    var dataToSerialize = new TaskFrequencyConfig
                    {
                        //ParameterItems = vm.ParameterITem.Cast<object>().ToList(),
                        ParameterItems = vm.ParameterITem.Cast<string>().ToList(),
                        SelectedEntry = vm.SelectedEntry,
                        MaxFrequency = vm.MaxFrequency,
                        MinFrequency = vm.MinFrequency,
                        Bands = vm.Bands,
                        Band1 = vm.Band1,
                        Band2 = vm.Band2,
                        Band3 = vm.Band3
                    };
                    taskItem.Conditions = JsonSerializer.Serialize(dataToSerialize);

                    var result = await taskHelper.AddTaskAsync(taskItem);

                    #region log record
                    var logRecord = new LogModel
                    {
                        TimeStamp = DateTime.Now,
                        UserName = Global.User.UserName,
                        TaskId = id.ToString(),
                        TaskName = vm.TaskName,
                        PN = null,
                        Level = vm.SelectedLevel,
                        SN = null,

                        ChipNumber = vm.SelectedPPTModel
                    };


                    if (result >= 0)
                    {
                        //添加日志
                        logRecord.Message += "Add Task Success!";
                        success = true;
                    }
                    else
                    {

                        logRecord.Message += "Add Task Failed!";
                        //添加日志
                        success = false;
                    }
                    logRecord.Message += JsonSerializer.Serialize(taskItem);

                    var result2 = await LogRepository.InsertLogAsync(logRecord);

                    if (result2)
                    {
                    }
                    else
                    {

                        System.Windows.MessageBox.Show("Insert Log Failed! Please Check the Database Connection!", "", MessageBoxButton.OK, MessageBoxImage.Error);

                    }
                    #endregion
                }
                catch (Exception ex)

                {
                    System.Windows.MessageBox.Show(ex.Message, "", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;

                }
            }
             return success;


        }
    }


    public class NewTaskWinModel : ObeservableObject
    {
        private string _parentObjectName;
        public string ParentObjectName
        {
            set { _parentObjectName=value; RaisePropertyChanged(nameof(ParentObjectName)); }
            get { return _parentObjectName; }
        }

        
        private DeviceTreeViewItem _deveiceTreeItem;
        public DeviceTreeViewItem DeviceTreeItem
        {
            get { return _deveiceTreeItem; }
            set { this._deveiceTreeItem = value; RaisePropertyChanged(nameof(DeviceTreeItem)); }
        }

        public ObservableCollection<DeviceTreeViewItem> DeviceTreeResources { get; set; } = new ObservableCollection<DeviceTreeViewItem>();
       
        private string _folderPath;
        public string FolderPath
        {
            set { _folderPath = value; RaisePropertyChanged(nameof(FolderPath)); }
            get { return _folderPath; }
        }

        private string _pptName;
        public string PptName
        {
            set { _pptName = value; RaisePropertyChanged(nameof(PptName)); }
                
                get { return _pptName; }
        }


        // 1. Task Name (TextBox)
        private string _taskName;
        public string TaskName
        {
            get { return _taskName; }
            set { _taskName = value; RaisePropertyChanged(nameof(TaskName)); /* 可以触发 CanExecuteChanged */ }
        }

        // 2. ComboBox / CheckComboBox Data Sources

        // Status ComboBox
        public ObservableCollection<string> StatusSource { get; set; } = new ObservableCollection<string>();

        // Level ComboBox
        public ObservableCollection<string> LevelSource { get; set; } = new ObservableCollection<string>();

        // Major Device ComboBox
        public ObservableCollection<string> MajorSource { get; set; } = new ObservableCollection<string>();

        // Minor Device ComboBox
        public ObservableCollection<string> MinorSource { get; set; } = new ObservableCollection<string>();

        // PPT Model ComboBox
        public ObservableCollection<string> PPTModelSource { get; set; } = new ObservableCollection<string>();

        // Parameters CheckComboBox (xctk:CheckComboBox)
        public ObservableCollection<string> ParametersSource { get; set; } = new ObservableCollection<string>();

        private ObservableCollection<object> _parameterItem = new ObservableCollection<object>();
        public ObservableCollection<object> ParameterITem
        {
            get { return _parameterItem; }
            set { _parameterItem = value; RaisePropertyChanged(); }
        }
        private string _selectedStatus;
        /// <summary> 绑定到 Status ComboBox 的 SelectedItem </summary>
        public string SelectedStatus
        {
            get { return _selectedStatus; }
            set { _selectedStatus = value; RaisePropertyChanged(nameof(SelectedStatus)); }
        }

        private string _selectedLevel;
        /// <summary> 绑定到 Level ComboBox 的 SelectedItem </summary>
        public string SelectedLevel
        {
            get { return _selectedLevel; }
            set { _selectedLevel = value; RaisePropertyChanged(nameof(SelectedLevel)); }
        }

        private string _selectedMajor;
        /// <summary> 绑定到 Device Major ComboBox 的 SelectedItem </summary>
        public string SelectedMajor
        {
            get { return _selectedMajor; }
            set { _selectedMajor = value; RaisePropertyChanged(nameof(SelectedMajor)); }
        }

        private string _selectedMinor;
        /// <summary> 绑定到 Device Minor ComboBox 的 SelectedItem </summary>
        public string SelectedMinor
        {
            get { return _selectedMinor; }
            set { _selectedMinor = value; RaisePropertyChanged(nameof(SelectedMinor)); }
        }

        private string _selectedPPTModel;
        /// <summary> 绑定到 PPT Model ComboBox 的 SelectedItem </summary>
        public string SelectedPPTModel
        {
            get { return _selectedPPTModel; }
            set { _selectedPPTModel = value; RaisePropertyChanged(nameof(SelectedPPTModel)); }
        }
        // 4. Second Part Properties

        // Selected Entrie (TextBox)
        private string _selectedEntry;
        public string SelectedEntry
        {
            get { return _selectedEntry; }
            set { _selectedEntry = value; RaisePropertyChanged(nameof(SelectedEntry)); }
        }

        // Min Frequency (xctk:DecimalUpDown)
        private double _minFrequency = 0.0;
        public double MinFrequency
        {
            get { return _minFrequency; }
            set { _minFrequency = value; RaisePropertyChanged(nameof(MinFrequency)); }
        }

        // Max Frequency (xctk:DecimalUpDown)
        private double _maxFrequency = 10.0;
        public double MaxFrequency
        {
            get { return _maxFrequency; }
            set { _maxFrequency = value; RaisePropertyChanged(nameof(MaxFrequency)); }
        }
        private string _bands;
        public string Bands
        {
            get { return _bands; }  
            set { _bands = value; RaisePropertyChanged(nameof(Bands)); }
        }
        public ObservableCollection<string> BandsSource { get; set; } = new ObservableCollection<string>() { "1","2","3"};
        // Frequency Band TextBoxes
        private double _band1MinValue;
        public double Band1MinValue
        {
            get { return _band1MinValue; }
            set { _band1MinValue = value; RaisePropertyChanged(nameof(Band1MinValue)); }
        }
        private double _band1MaxValue;
        public double Band1MaxValue
        {
            get { return _band1MaxValue; }
            set { _band1MaxValue = value; RaisePropertyChanged(nameof(Band1MaxValue)); }
        }




        private string _band1;
        public string Band1
        {
            get { return _band1; }
            set { _band1 = value; RaisePropertyChanged(nameof(Band1)); }
        }


        private double _band2MinValue;
        public double Band2MinValue
        {
            get { return _band2MinValue; }
            set { _band2MinValue = value; RaisePropertyChanged(nameof(Band2MinValue)); }
        }
        private double _band2MaxValue;
        public double Band2MaxValue
        {
            get { return _band2MaxValue; }
            set { _band2MaxValue = value; RaisePropertyChanged(nameof(Band2MaxValue)); }
        }
        private string _band2;
        public string Band2
        {
            get { return _band2; }
            set { _band2 = value; RaisePropertyChanged(nameof(Band2)); }
        }

        // 注意: Band3 在 XAML 中默认是 Collapsed 的

        private double _band3MinValue;
        public double Band3MinValue
        {
            get { return _band3MinValue; }
            set { _band3MinValue = value; RaisePropertyChanged(nameof(Band3MinValue)); }
        }
        private double _band3MaxValue;
        public double Band3MaxValue
        {
            get { return _band3MaxValue; }
            set { _band3MaxValue = value; RaisePropertyChanged(nameof(Band3MaxValue)); }
        }

        private string _band3;
        public string Band3
        {
            get { return _band3; }
            set { _band3 = value; RaisePropertyChanged(nameof(Band3)); }
        }
    }


    public class DeviceTreeViewItem : ObeservableObject
    {
        private string content;
        public string Content
        {
            get { return content; }
            set { this.content = value; RaisePropertyChanged(nameof(content)); }
        }
        public DeviceTreeViewItem Parent { get; set; }

        public bool IsLeaf
        {
            // 如果 Children 集合为空（或 Count 为 0），则为叶子节点
            get { return Children == null || Children.Count == 0; }
        }

        private bool visible = false;
        public bool Visible
        {
            get { return visible; }
            set { this.visible = value; RaisePropertyChanged(nameof(visible)); }
        }
        public ObservableCollection<DeviceTreeViewItem> Children { get; set; } = new ObservableCollection<DeviceTreeViewItem>();

        public DeviceTreeViewItem()
        {
            Children.CollectionChanged += (sender, e) => {
                // 当子集合变化时，通知绑定系统 IsLeaf 属性可能已改变
                RaisePropertyChanged(nameof(IsLeaf));
            };
        }
    }


    public class TaskSources
    {
        public List<string> StatusSource { get; set; }
        public List<string> LevelSource { get; set; }
        public List<string> MajorSource { get; set; }
        public List<string> MinorSource { get; set; }
        public List<string> PPTModelSource { get; set; }
        //public List<string> ParametersSource { get; set; }
        public List<string> BandsSource { get; set; }
    }

    public class TaskDefaultValues
    {
        public string TaskName { get; set; }
        public string FolderPath { get; set; }

        public string SelectedStatus { get; set; }
        public string SelectedLevel { get; set; }
        public string SelectedMajor { get; set; }
        public string SelectedMinor { get; set; }
        public string SelectedPPTModel { get; set; }
        public string SelectedEntry { get; set; }

        // 注意：这里使用 double 匹配您的 ViewModel 属性定义
        public double MinFrequency { get; set; }
        public double MaxFrequency { get; set; }

        public string Band1 { get; set; }
        public string Band2 { get; set; }
        public string Band3 { get; set; }
    }

    public class InitialTaskData
    {
        public TaskSources Sources { get; set; }
        public TaskDefaultValues DefaultValues { get; set; }
    }

    public class RootObject
    {
        //映射到 JSON 字符串的属性名称， 用于反序列化 来自 System.Text.Json.Serialization
        [JsonPropertyName("InitialData")]
        public InitialTaskData InitialData { get; set; }
    }

    // 专门用于序列化的数据传输对象
    public class TaskFrequencyConfig
    {
        public string Band1 { get; set; }
        public string Band2 { get; set; }
        public string Band3 { get; set; }
        public double MaxFrequency { get; set; }
        public double MinFrequency { get; set; }
        public List<string> ParameterItems { get; set; } // 注意类型与集合保持一致
        public string SelectedEntry { get; set; }
        public string Bands { get; set; } // 假设你的 Bands 属性是 string
    }

}
