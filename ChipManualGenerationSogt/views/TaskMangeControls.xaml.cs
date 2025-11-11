using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using CommunityToolkit.Mvvm.Input;
using ScottPlot.Finance;
using System.Windows.Controls.Primitives;
using Microsoft.Web.WebView2.Core;
using DocumentFormat.OpenXml.Office2021.DocumentTasks;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Presentation;
using Xceed.Wpf.Toolkit;
using CommunityToolkit.Mvvm.ComponentModel;
using System.Text.Json;
using DocumentFormat.OpenXml.Wordprocessing;
namespace ChipManualGenerationSogt
{
    /// <summary>
    /// TaskMangeControls.xaml 的交互逻辑
    /// </summary>
    public partial class TaskMangeControls : System.Windows.Controls.UserControl
    {
        //const string K_STATUS_1 = "Finished"; //1 
        //const string K_STATUS_2 = "In Progress";//2
        //const string K_STATUS_3 = "Not Started";//3 
        public event EventHandler<TaskTableItem> TaskExcute;
        public event EventHandler    TestEvent;
        public event EventHandler    AddEvent;
        public event EventHandler    QueryLogEvent;
        public event EventHandler<TaskTableItem> DetailShowEvent;
        TaskMangeControlsViewModel vm;
        User _currentUser;
        public TaskMangeControls()
        {
            InitializeComponent();
            vm = new TaskMangeControlsViewModel();
            this.DataContext = vm;
            init();
            _currentUser = new User();


        }

        private async void init()
        {

            //string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
            //var repo = new TaskRepository(connStr);
            //var taskItems = await repo.GetAllCurrentTasksAsync();
            //foreach (var item in taskItems)
            //{
            //    var uiItem = DBTaskToTaskModel(item);
            //    uiItem.TaskFinish += OnTaskFinshed;
            //    uiItem.PopupShow += OnPopupShow;
            //    uiItem.Ok += OnOk;

            //    vm.CurrentTasks.Add(uiItem);

            //}

            //vm.PopupTopMessage = "A: 这是一个测试.\nB: 这是一个测试回复.\nC: 这是一个测试回复.";
            //vm.PopupTitle = "增加测试记录任务的消息提示框";

            //var sqlServer = new TaskSqlServerRepository();
            //var taskItems = await sqlServer.GetAllTasksAsync();
            //foreach (var item in taskItems)
            //{
            //    //var uiItem = DBTaskToTaskModel(item);
            //    //uiItem.TaskFinish += OnTaskFinshed;
            //    //uiItem.PopupShow += OnPopupShow;
            //    //uiItem.Ok += OnOk;

            //    //vm.CurrentTasks.Add(uiItem);
            //    vm.TableItemSources.Add(DBTaskToUITaskModel(item));
            //}
            RefreshTask();


        }
        private TaskTableItem DBTaskToUITaskModel(TaskSqlServerModel item)
        {
            var uiItem = new TaskTableItem
            {
                ID = item.ID,
                PPTModel = item.PPTModel,
                TaskName = item.TaskName,
                PptName = item.PptName,
                Status = item.Status,
                Level = item.Level.Trim(),
                DataStatus = item.DataStatus,
                FilesStatus = item.FilesStatus,
                StartDate = item.StartDate,
                EndDate = item.EndDate,
                Conditions = item.Conditions,
                Major = item.Major,
                Minor = item.Minor,
                TableUpdate = item.TableUpdate,
            };
            if (item.Status == "PPT Ready For Generate")
            {
                uiItem.MenuItemCommitIsEnabled = false;
                uiItem.MenuItemBackToModifyIsEnabled = true;
            }

            else if (item.Status == "Not Commited")
            {
                uiItem.MenuItemCommitIsEnabled = true;
                uiItem.MenuItemBackToModifyIsEnabled = false;
            } else if(item.Status == "Completed")
            {

                uiItem.MenuItemCommitIsEnabled = false;
                uiItem.MenuItemBackToModifyIsEnabled = false;
            }

                return uiItem;
        }


        private void DataGrid_Selected(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("hello");
        }

        private async void OnPptPreview(object _unusedSender, TaskTableItem task)
        {
            try
            {
                string ftpPath = $"{task.Major}\\{(task.Minor ?? "")}\\{task.TaskName}";
                vm.IsBusy = true;
                string filePath = System.IO.Path.Combine(ftpPath, task.PptName);
                string localPath = System.IO.Path.Combine(Global.TempBasePath, task.PptName);
                if (await FtpClient.DownloadFileAsyncEx(filePath, localPath))
                {
                    string fileName = System.IO.Path.GetFileName(filePath);
                    string pptFile = System.IO.Path.Combine(Global.TempBasePath, fileName);
                    string pdfFile = System.IO.Path.Combine(Global.TempBasePath, System.IO.Path.GetFileNameWithoutExtension(fileName) + ".pdf");
                    await System.Threading.Tasks.Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));
                    var pdfShown = new PdfShowWin();
                    pdfShown.Status = true;
                    PdfShowWin.PPTPath = pptFile;
                    PdfShowWin.PdfPath = pdfFile;
                    pdfShown.ShowPdf(pdfFile);
                    pdfShown.Show();
                }else
                {
                    System.Windows.MessageBox.Show("Download failed.", "", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message,"", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            { 
                vm.IsBusy = false;
            }



        }
        private void OnDetaileEvent(object sender, TaskTableItem task)
        {
            //vm.PopupIsOpen = true;
            //vm.PopupTitle = task.TaskName;
            
            //vm.PopupTopMessage = FormatJson(task.Conditions);

            DetailShowEvent?.Invoke(sender, task);
        }

        private async void OnDownloadEvent(object _unusedSender, TaskTableItem task)
        {

            try
            {
                using (var dialog = new FolderBrowserDialog() )
                {
                    dialog.SelectedPath = Global.AppBaseUrl;
                    dialog.Description = "Select the folder to save the files";
                    dialog.ShowNewFolderButton = true;
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                       
                        vm.IsBusy = true;
                        string ftpPath = $"{task.Major}\\{(task.Minor ?? "")}\\{task.TaskName}";
                      
                        string filePath = System.IO.Path.Combine(ftpPath, task.PptName);
                        string localPath = System.IO.Path.Combine(dialog.SelectedPath, task.PptName);
                        if (await FtpClient.DownloadFileAsyncEx(filePath, localPath))
                        {
                            System.Windows.MessageBox.Show("Download Success.", "", MessageBoxButton.OK, MessageBoxImage.Information);

                        }
                        else
                        {
                            System.Windows.MessageBox.Show("Download failed.", "", MessageBoxButton.OK, MessageBoxImage.Error);
                        }

                    }
                }
               
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                vm.IsBusy = false;
            }

        }


        private async void OnCommitedEvent(object _unusedSender, TaskTableItem task)
        {
            var resoult =System.Windows.MessageBox.Show("Are You Sure to Commit the Task?", "Tips", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (resoult == MessageBoxResult.OK)
            { 
                if((UserPriority)Global.User.priority == UserPriority.DataProvider)
                {
                    task.Status = "PPT Ready For Generate";
                    task.MenuItemCommitIsEnabled = false;
                }
                
                
            }
        
        }

        private async void OnModifiedEvent(object _unusedSender, TaskTableItem task)
        {

            var resoult = System.Windows.MessageBox.Show("Are You Sure to Back To Modify The Task Data Source?", "Tips", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (resoult == MessageBoxResult.OK)
            {
                if ((UserPriority)Global.User.priority == UserPriority.PptMaker)
                {
                    task.Status = "PPT Ready For Generate";
                    task.MenuItemCommitIsEnabled = false;
                }


            }

        }
        private void doubleClick(object sender, MouseButtonEventArgs e)
        {
            Console.WriteLine("hello world");
            // 确保点击的是行（而不是标题、空白区域）
            var row = FindAncestor<DataGridRow>((DependencyObject)e.OriginalSource);
            if (row == null) return;

            // 获取当前选中的数据项（即你的业务对象，如 TaskItem）
            var selectedItem = dataGrid.SelectedItem;
            if (selectedItem == null) return;

            // 执行你的逻辑，例如：
            var task = selectedItem as TaskTableItem; // 替换为你的实际类型
            //MessageBox.Show($"双击了任务：{task.TaskName}");
            Console.WriteLine($"{task.TaskName}");
            Global.TaskModel = task;
            Global.FtpRootPath = $"{task.Major}\\{(task.Minor ?? "")}\\{task.TaskName}";
            //Global.TaskModel = await sqlServer.GetTaskByIdAsync();

            TaskExcute?.Invoke(this, task);

        }

        public static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T ancestor)
                    return ancestor;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        public string FormatJson(string compactJsonString)
        {
            try
            {
                // 1. 反序列化: 将紧凑的 JSON 字符串解析成一个 JSON Document (可以理解为一个内存中的对象结构)
                using (JsonDocument document = JsonDocument.Parse(compactJsonString))
                {
                    // 2. 序列化选项: 启用缩进格式 (WriteIndented = true)
                    var options = new JsonSerializerOptions { WriteIndented = true };

                    // 3. 重新序列化: 将内存中的对象结构格式化成带缩进的字符串
                    return JsonSerializer.Serialize(document.RootElement, options);
                }
            }
            catch (JsonException)
            {
                // 如果 JSON 字符串格式错误，返回原始字符串或错误信息
                return "JSON 格式错误，无法美化。";
            }
        }
        // 这是之前用于增加的按钮事件
        //private async void Add_btn_Clicked(object sender, RoutedEventArgs e)
        //{
        //    if (int.TryParse(vm.NewTask.ID, out int id))
        //    {

        //    }
        //    else 
        //    {
        //        System.Windows.MessageBox.Show("The input is not a valid integer","", MessageBoxButton.OK, MessageBoxImage.Warning);
        //        return;
        //    }

        //    var result = System.Windows.MessageBox.Show(
        //        $"Are you sure to add the task: {vm.NewTask.TaskName}?",
        //        "Confirm Add", // 建议加上标题
        //        MessageBoxButton.OKCancel, // ? 用 OKCancel，不是 OK
        //        MessageBoxImage.Question
        //    );

        //    if (result == MessageBoxResult.OK)
        //    {




        //        // 先写入数据库

        //        var status = vm.NewTask.Status == K_STATUS_1 ? 1 :
        //                     vm.NewTask.Status == K_STATUS_2 ? 2 :
        //                     vm.NewTask.Status == K_STATUS_3 ? 3 : 0;

        //        int idtmp = Convert.ToInt32(vm.NewTask.ID);
        //        var task = new TaskModel_DB
        //        {
        //            ID = idtmp,
        //            TaskName = vm.NewTask.TaskName,
        //            Status = 3,              // 假设 1=进行中，0=未开始
        //            StartDate = DateTime.Parse(vm.NewTask.StartDate),
        //            Consumed = 0,              // 已消耗时间（分钟？）
        //            DataStatus = false,          // 数据已就绪
        //            FilesStatus = false,         // 文件已下载
        //            CheckStatus = false         // 尚未审核
        //        };
        //        string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
        //        var repo = new TaskRepository(connStr);
        //        var index = await repo.InsertCurrentTaskAsync(task);
        //        if (index > 0)
        //        {
        //            var uiItem= DBTaskToTaskModel(task);
        //            uiItem.TaskFinish += OnTaskFinshed;
        //            uiItem.PopupShow += OnPopupShow;
        //            uiItem.Ok += OnOk;
        //            vm.CurrentTasks.Add(uiItem);

        //        }
        //        var sqlServer = new TaskRepository();
        //        var opModel = new OperationModel
        //        {
        //            TaskID = idtmp,
        //            TaskName = vm.NewTask.TaskName,
        //            TimeStamp = DateTime.Now,
        //        };
        //        await sqlServer.InsertOperationAsync(opModel);

        //        var logModel = new LogModel
        //        {
        //            TimeStamp =DateTime.Now,
        //            UserName = vm.NewTask.TaskName,
        //            Level = LogLevels.Info,
        //            TaskId = vm.NewTask.ID,
        //            TaskName = vm.NewTask.TaskName,
        //            Message = $"The task {vm.NewTask.TaskName} has been generate by {vm.NewTask.TaskName}."

        //        };

        //        await sqlServer.InsertLogAsync(logModel);

        //    }


        //}




        private void ClosePopup(object sender, RoutedEventArgs e)
        {
            vm.PopupIsOpen = false;
        }
   
        public void SetCurrentUser(User user)
        {
            _currentUser = user;
        }





        private async void Btn_Test_Click(object sender, RoutedEventArgs e)
        {


            string ftpRootPath = "Amplifier/MML806";

            //string folderPath = "F:\\PROJECT\\ChipManualGeneration\\原始数据\\MML004X 手册数据-3";
            //await FtpClient.UploadFolderAsync(folderPath, ftpRootPath);

            //await FtpClient.DownloadFolderAsync(ftpRootPath, Global.FileBasePath);

            var pdfWindow = new PdfShowWin();
            pdfWindow.ShowPdf("F:\\PROJECT\\ChipManualGeneration\\exe\\app\\ChipManualGenerationSogt\\bin\\Debug\\resources\\files\\demo.pdf");
            pdfWindow.Show();

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

        // 辅助方法：向上查找可视化树中的父控件(你需要将这个方法添加到你的类中)
            // 它应该与你之前的 FindVisualChild 放在同一个类中。
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

        private void Btn_Add_Click(object sender, RoutedEventArgs e)
        {
            AddEvent?.Invoke(this, EventArgs.Empty);
        }

        public async void RefreshTask()
        {
            vm.TableItemSources.Clear();
            var sqlServer = new TaskSqlServerRepository();
            var taskItems = await sqlServer.GetAllTasksAsync();
            
            foreach (var item in taskItems)
            {
                var uiItem = DBTaskToUITaskModel(item);
                uiItem.PreviewEvent += OnPptPreview;
                uiItem.DetaileEvent += OnDetaileEvent;
                uiItem.DownloadEvent += OnDownloadEvent;
                uiItem.CommitEvent += OnCommitedEvent;
                uiItem.ModifyEvent += OnModifiedEvent;
                //uiItem.MenuItemBackToModifyIsEnabled = false;
                //uiItem.MenuItemCommitIsEnabled = true;
                vm.TableItemSources.Add(uiItem);
            }
        }

        private void Btn_QueryLog_Click(object sender, RoutedEventArgs e)
        {
            QueryLogEvent?.Invoke(this, EventArgs.Empty);
        }

        private void MenuItem_Commit_Click(object sender, RoutedEventArgs e)
        {

        }

        private void MenuItem_BackToModify_Click(object sender, RoutedEventArgs e)
        {

        }


    }

    public class TaskMangeControlsViewModel : ObeservableObject
    {
        private bool _popupIsOpen;
        public bool PopupIsOpen { get => _popupIsOpen; set { _popupIsOpen = value; RaisePropertyChanged(); } }

        private string _popupTopMessage;
        public string PopupTopMessage { get => _popupTopMessage; set { _popupTopMessage = value; RaisePropertyChanged(); } }

        private string _popupBottomMessage;
        public string PopupBottomMessage { get => _popupBottomMessage; set { _popupBottomMessage = value; RaisePropertyChanged(); } }

        private string _popupTitle;
        public string PopupTitle { get => _popupTitle; set { _popupTitle = value; RaisePropertyChanged(); } }
        public TaskMangeControlsViewModel()
        {
           
        }
        //public ObservableCollection<TaskModel> CurrentTasks { get; set; } = new ObservableCollection<TaskModel>();

        //public ObservableCollection<TaskModel> FinishedTasks { get; set; } = new ObservableCollection<TaskModel>();

        //private TaskModel _newTask = new TaskModel();
        //public TaskModel NewTask { get => _newTask; set { _newTask = value; RaisePropertyChanged(); } }

        private ObservableCollection<object> _selectedLevels = new ObservableCollection<object>();
        public ObservableCollection<object> SelectedLevels
        {
            get { return _selectedLevels; }
            set
            {
                _selectedLevels = value;
                // 确保你的 RaisePropertyChanged 方法存在
                RaisePropertyChanged();
            }
        }

        public ObservableCollection<string>   Levels { get; set; } = new ObservableCollection<string>() { "Low", "Medium", "High" };


        private ObservableCollection<object> _selectedStatus = new ObservableCollection<object>();
        public ObservableCollection<object> SelectedStatus
        {
            get { return _selectedStatus; }
            set
            {
                _selectedStatus = value;
                // 确保你的 RaisePropertyChanged 方法存在
                RaisePropertyChanged();
            }
        }
        public ObservableCollection<string> StatusSources { get; set; } = new ObservableCollection<string>() { "In Progress", "Completed"};


        private TaskTableItem _selectTableItem;
        public TaskTableItem SelectTableItem
        {
            get { return _selectTableItem; }
            set
            {
                _selectTableItem = value;
                // 确保你的 RaisePropertyChanged 方法存在
                RaisePropertyChanged();
            }
        }

        public ObservableCollection<TaskTableItem> TableItemSources { get; set; } = new ObservableCollection<TaskTableItem>();
        // 新的表格

        bool _isBusy;
        public bool IsBusy
        {
            get { return _isBusy; }
            set { _isBusy = value; RaisePropertyChanged(); }
        }

        string _busyMessage;
        public string BusyMessage
        {
            get { return _busyMessage; }
            set { _busyMessage = value; RaisePropertyChanged(); }
        }

        public event EventHandler AddEvent;
        //public event EventHandler<TaskModel> DetaileEvent;

        //public void OnDetaileEvent(object sender, TaskModel e)
        //{
        //    DetaileEvent?.Invoke(sender, e);
        //}

        public void OnAddEvent(object sender, EventArgs e)
        {
            AddEvent?.Invoke(sender, e);
        }
    }


    // 假设您使用的是 Community Toolkit MVVM
// using CommunityToolkit.Mvvm.ComponentModel; 


    public class TaskTableItem : ObservableObject
    {
       
        public TaskSqlServerModel DataModel { get; private set; }

        
        public TaskTableItem(TaskSqlServerModel model = null)
        {
            // 如果提供了模型，则使用它；否则创建新实例
            DataModel = model ?? new TaskSqlServerModel();
            DetailShowCommand = new RelayCommand<TaskTableItem>(OnShowPopup);
            PreviewCommand = new RelayCommand<TaskTableItem>(OnPreview);
            DownLoadCommand = new RelayCommand<TaskTableItem>(OnDownload);

            CommitCommand = new RelayCommand<TaskTableItem>(OnCommit);
            ModifyCommand = new RelayCommand<TaskTableItem>(OnModify);
        }

        // --- 封装所有需要通知 UI 的属性 ---


        public int ID
        {
            get => DataModel.ID;
            set => SetProperty(DataModel.ID, value, DataModel, (model, val) => model.ID = val);
        }

        // 2. PPTModel
        public string PPTModel
        {
            get => DataModel.PPTModel;
            set => SetProperty(DataModel.PPTModel, value, DataModel, (model, val) => model.PPTModel = val);
            // 或者更简洁的标准方法：
            // set
            // {
            //     if (DataModel.PPTModel != value)
            //     {
            //         DataModel.PPTModel = value;
            //         OnPropertyChanged(); // 通知 UI
            //     }
            // }
        }

        // 3. TaskName
        public string TaskName
        {
            get => DataModel.TaskName;
            set
            {
                if (SetProperty(DataModel.TaskName, value, (val) => DataModel.TaskName = val))
                {
                    
                }
            }
        }

        // 4. Status
        public string Status
        {
            get => DataModel.Status;
            set => SetProperty(DataModel.Status, value, DataModel, (model, val) => model.Status = val);
        }

        // 5. Level
        public string Level
        {
            get => DataModel.Level;
            set => SetProperty(DataModel.Level, value, DataModel, (model, val) => model.Level = val.Trim());
        }

        // 6. Major
        public string Major
        {
            get => DataModel.Major;
            set => SetProperty(DataModel.Major, value, DataModel, (model, val) => model.Major = val);
        }

        // 7. Minor
        public string Minor
        {
            get => DataModel.Minor;
            set => SetProperty(DataModel.Minor, value, DataModel, (model, val) => model.Minor = val);
        }

        // 8. StartDate
        public DateTime StartDate
        {
            get => DataModel.StartDate;
            set => SetProperty(DataModel.StartDate, value, DataModel, (model, val) => model.StartDate = val);
        }

        // 9. EndDate
        public DateTime? EndDate
        {
            get => DataModel.EndDate;
            set => SetProperty(DataModel.EndDate, value, DataModel, (model, val) => model.EndDate = val);
        }

        // 10. DataStatus
        public bool DataStatus
        {
            get => DataModel.DataStatus;
            
            set => SetProperty(DataModel.DataStatus, value, DataModel, (model, val) => model.DataStatus = val);
        }

        // 11. FilesStatus
        public bool FilesStatus
        {
            get => DataModel.FilesStatus;
            set => SetProperty(DataModel.FilesStatus, value, DataModel, (model, val) => model.FilesStatus = val);
        }

        // 12. Conditions
        public string Conditions
        {
            get => DataModel.Conditions;
            set => SetProperty(DataModel.Conditions, value, DataModel, (model, val) => model.Conditions = val);
        }

        public event EventHandler<TaskTableItem> DetaileEvent;

        public ICommand DetailShowCommand { get; }
        

        [RelayCommand]
        private void OnShowPopup(TaskTableItem item)
        {
            DetaileEvent?.Invoke(this, item);

        }
    
        public string PptName
        {
            get { return DataModel.PptName; }
            set => SetProperty(DataModel.PptName, value, DataModel, (model, val) => model.PptName = val);
        }
        public ICommand PreviewCommand { get; }
        public event EventHandler<TaskTableItem> PreviewEvent;
        private  void OnPreview(TaskTableItem item)
        {
            if (Status == "Completed")
            {
                PreviewEvent? .Invoke(this,item ); 
            }
            else
            {
                System.Windows.MessageBox.Show("Please finish the task first!");
            }
        }

        public ICommand DownLoadCommand { get; }
        public event EventHandler<TaskTableItem> DownloadEvent;
        private void OnDownload(TaskTableItem item)
        {
            if (Status == "Completed")
            {
                DownloadEvent?.Invoke(this, item);
            }
            else
            {
                System.Windows.MessageBox.Show("Please finish the task first!");
            }
        }
       
        public bool? TableUpdate { get; set; } 


        public ICommand CommitCommand { get; }

        public event EventHandler<TaskTableItem> CommitEvent;

        private void OnCommit(TaskTableItem item)
        {
            CommitEvent?.Invoke(this, item);
        }

        public ICommand ModifyCommand { get; }

        public event EventHandler<TaskTableItem> ModifyEvent;
        private void OnModify(TaskTableItem item)
        {
            ModifyEvent?.Invoke(this, item);
        }

        public bool MenuItemCommitIsEnabled { get; set; }
        public bool MenuItemBackToModifyIsEnabled { get; set; }
    }
}
