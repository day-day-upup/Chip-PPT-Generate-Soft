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
namespace ChipManualGenerationSogt
{
    /// <summary>
    /// TaskMangeControls.xaml 的交互逻辑
    /// </summary>
    public partial class TaskMangeControls : System.Windows.Controls.UserControl
    {
        const string K_STATUS_1 = "Finished"; //1 
        const string K_STATUS_2 = "In Progress";//2
        const string K_STATUS_3 = "Not Started";//3 
        public event EventHandler<TaskModel> TaskExcute;
        public event EventHandler    TestEvent;
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

            string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
            var repo = new TaskRepository(connStr);
            var taskItems = await repo.GetAllCurrentTasksAsync();
            foreach (var item in taskItems)
            {
                var uiItem = DBTaskToTaskModel(item);
                uiItem.TaskFinish += OnTaskFinshed;
                uiItem.PopupShow += OnPopupShow;
                uiItem.Ok += OnOk;

                vm.CurrentTasks.Add(uiItem);
            
            }

            vm.PopupTopMessage = "A: 这是一个测试.\nB: 这是一个测试回复.\nC: 这是一个测试回复.";
            vm.PopupTitle = "增加测试记录任务的消息提示框";
        }
        private void DataGrid_Selected(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("hello");
        }
        private async void OnTaskFinshed1(object sender,TaskModel task)
        {
            string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
            var repo = new TaskRepository(connStr);
            var taskItems = await repo.GetAllCurrentTasksAsync();
            var taskItem = taskItems.FirstOrDefault(t => t.ID == Convert.ToInt32(task.ID));
            Console.WriteLine("任务完成");
            if (_currentUser.priority == 0)//管理员
            {
                
            }
            else if (_currentUser.priority == 1) // 提供数据者
            {
                taskItem.DataStatus = true;
                if (await repo.UpdateCurrentTaskAsync(taskItem))
                {

                    TaskModel taskToUpdate = vm.CurrentTasks
                            .FirstOrDefault(t => t.ID == task.ID);

                    if (taskToUpdate != null)
                    {
                        taskToUpdate.DataStatus = "True";
                    }
                }

                var sqlServer = new TaskRepository();
                var logModel = new LogModel
                {
                    TaskId = task.ID,
                    UserName = Global.User.UserName,
                    TimeStamp = DateTime.Now,
                    TaskName = task.TaskName,
                    Level = LogLevels.Error,
                    Message = $"The  data status of the  {task.TaskName}  is finished."
                };
                if (!await sqlServer.InsertLogAsync(logModel))
                {
                    System.Windows.MessageBox.Show("Add log failed., please check the connection  of data base", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
            else if (_currentUser.priority == 2) //ppt 制作者
            {
                if (taskItem.DataStatus)
                {
                    taskItem.FilesStatus = true;

                    if (await repo.UpdateCurrentTaskAsync(taskItem))
                    {

                        TaskModel taskToUpdate = vm.CurrentTasks
                                .FirstOrDefault(t => t.ID == task.ID);

                        if (taskToUpdate != null)
                        {
                            taskToUpdate.FilesStatus = " True";
                        }
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Data not ready, please check the data status first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                var sqlServer = new TaskRepository();
                var logModel = new LogModel
                {
                    TaskId = task.ID,
                    UserName = Global.User.UserName,
                    TimeStamp = DateTime.Now,
                    TaskName = task.TaskName,
                    Level = LogLevels.Error,
                    Message = $"The  file status of the  {task.TaskName}  is finished."
                };
                if (!await sqlServer.InsertLogAsync(logModel))
                {
                    System.Windows.MessageBox.Show("Add log failed., please check the connection  of data base", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
            else if (_currentUser.priority == 3) // 审核者
            {

                if (taskItem.DataStatus && taskItem.FilesStatus)
                {
                    taskItem.CheckStatus = true;
                    taskItem.Status = 1;
                    //if (await repo.UpdateFinishedTaskAsync(taskItem) && await repo.DeleteCurrentTaskAsync(taskItem))

                        if (await repo.InsertFinishedTaskAsync(taskItem) && await repo.DeleteCurrentTaskAsync(taskItem))
                    {

                        TaskModel taskToUpdate = vm.CurrentTasks
                                .FirstOrDefault(t => t.ID == task.ID);

                        if (taskToUpdate != null)
                        {
                            taskToUpdate.FilesStatus = "True";
                            //taskToUpdate.CheckStatus = "True";
                            taskToUpdate.Status = K_STATUS_1;
                        }
                        vm.CurrentTasks.Remove(task);
                        vm.FinishedTasks.Add(task);
                    }

                    var sqlServer = new TaskRepository();
                    var logModel = new LogModel
                    {
                        TaskId = task.ID,
                        UserName = Global.User.UserName,
                        TimeStamp = DateTime.Now,
                        TaskName = task.TaskName,
                        Level = LogLevels.Error,
                        Message = $"The  check status of the  {task.TaskName}  is finished."
                    };
                    if (!await sqlServer.InsertLogAsync(logModel))
                    {
                        System.Windows.MessageBox.Show("Add log failed., please check the connection  of data base", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }

                }
                else
                {
                    System.Windows.MessageBox.Show("Data  and file not ready, please check the data status first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    
                }
                //vm.CurrentTasks.Remove(task);
                //vm.FinishedTasks.Add(task);
            }

            

  
            
        
        }

        private async void OnTaskFinshed(object sender, TaskModel task)
        {
            string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
            var repo = new TaskRepository(connStr);

            // 性能优化：尽量减少数据库查询
            // 更好的做法是，TaskModel task 参数中已经包含了所有需要的信息，避免再次查询 GetAllCurrentTasksAsync
            var taskItems = await repo.GetAllCurrentTasksAsync();
            var taskItem = taskItems.FirstOrDefault(t => t.ID == Convert.ToInt32(task.ID));

            if (taskItem == null)
            {
                Console.WriteLine($"错误：未找到 ID 为 {task.ID} 的任务。");
                return;
            }

            Console.WriteLine("任务完成");

            // 使用枚举或常量增加可读性
            if (_currentUser.priority == (int)UserPriority.Admin)
            {
                // 管理员逻辑...
            }
            else if (_currentUser.priority == (int)UserPriority.DataProvider)
            {
                await HandleDataProviderFinish(repo, task, taskItem);
            }
            else if (_currentUser.priority == (int)UserPriority.PptMaker)
            {
                await HandlePptMakerFinish(repo, task, taskItem);
            }
            else if (_currentUser.priority == (int)UserPriority.Reviewer)
            {
                await HandleReviewerFinish(repo, task, taskItem);
            }
        }
        private void OnPopupShow(object sender, EventArgs e)
        {
            vm.PopupIsOpen = true;
        }

        async private void  OnOk(object sender, TaskModel task)
        {
            string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
            var repo = new TaskRepository(connStr);
            var taskItems = await repo.GetAllCurrentTasksAsync();
            var taskItem = taskItems.FirstOrDefault(t => t.ID ==Convert.ToInt32( task.ID));
            if (taskItem.Status == 3)
            {
                taskItem.Status = 2;
                if (await repo.UpdateCurrentTaskAsync(taskItem))
                {

                    TaskModel taskToUpdate = vm.CurrentTasks
                            .FirstOrDefault(t => t.ID == task.ID);

                    if (taskToUpdate != null)
                    {
                        taskToUpdate.Status = K_STATUS_2;
                    }
                }
                var sqlServer = new TaskRepository();
                var logModel = new LogModel
                {
                    TaskId = task.ID,
                    UserName =Global.User.UserName,
                    TimeStamp = DateTime.Now,
                    TaskName = task.TaskName,
                    Level = LogLevels.Info,
                    Message =$"The status of the task {task.TaskName} has been changed to In Progress."
                };
                if (!await sqlServer.InsertLogAsync(logModel))
                {
                    System.Windows.MessageBox.Show("Add log failed., please check the connection  of data base", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
               

                //“MessageBox”是“System.Windows.Forms.MessageBox”和“System.Windows.MessageBox”之间的不明确的引用
            }
            else {

                System.Windows.MessageBox.Show("The task change status failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                var sqlServer = new TaskRepository();
                var logModel = new LogModel
                {
                    TaskId = task.ID,
                    UserName = Global.User.UserName,
                    TimeStamp = DateTime.Now,
                    TaskName = task.TaskName,
                    Level = LogLevels.Error,
                    Message = $"The status of the task {task.TaskName} has changed failed"
                };
                if (!await sqlServer.InsertLogAsync(logModel))
                {
                    System.Windows.MessageBox.Show("Add log failed., please check the connection  of data base", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
            var task = selectedItem as TaskModel; // 替换为你的实际类型
            //MessageBox.Show($"双击了任务：{task.TaskName}");
            Console.WriteLine($"{task.TaskName}");
            Global.TaskModel = task;
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

        private  async void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //“TabControl”是“System.Windows.Controls.TabControl”和“System.Windows.Forms.TabControl”之间的不明确的引用
            var senderControl = sender as System.Windows.Controls.TabControl;
            if (senderControl == null || e.AddedItems.Count == 0)
            {
                return; // 如果不是选中操作，或 sender 不是 TabControl，则退出
            }
            if (senderControl != null)
            {
                // 方法1：通过 SelectedItem
                //if (tabControl.SelectedItem is TabItem selectedTab)
                //{
                //    Console.WriteLine($"选中的 Tab: {selectedTab.Header} (Name: {selectedTab.Name})");
                //}

                // 方法2：通过 SelectedIndex
                int index = tabControl.SelectedIndex;
                if (index == 1)
                {
                    vm.FinishedTasks.Clear();
                    string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
                    var repo = new TaskRepository(connStr);
                    var taskItems =await repo.GetAllFinishedTasksAsync();
                    foreach (var item in taskItems)
                    {
                        var uiItem = DBTaskToTaskModel(item);
                       
                        vm.FinishedTasks.Add(DBTaskToTaskModel(item));

                    }
                }
                if (index == 2)
                {
                    try
                    {
                        string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
                        var repo = new TaskRepository(connStr);
                        var lastTask = await repo.GetLastCurrentTasksAsync();
                        if (lastTask == null)
                            vm.NewTask.ID = "1";
                        else
                            vm.NewTask.ID = Convert.ToString(lastTask.ID + 1);
                        vm.NewTask.StartDate = DateTime.Now.ToString("yyyy.MM.dd hh:mm:ss");
                        int x = 1;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                
            }
        }

        private async void Add_btn_Clicked(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(vm.NewTask.ID, out int id))
            {

            }
            else 
            {
                System.Windows.MessageBox.Show("The input is not a valid integer","", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = System.Windows.MessageBox.Show(
                $"Are you sure to add the task: {vm.NewTask.TaskName}?",
                "Confirm Add", // 建议加上标题
                MessageBoxButton.OKCancel, // ? 用 OKCancel，不是 OK
                MessageBoxImage.Question
            );

            if (result == MessageBoxResult.OK)
            {
              

                

                // 先写入数据库
               
                var status = vm.NewTask.Status == K_STATUS_1 ? 1 :
                             vm.NewTask.Status == K_STATUS_2 ? 2 :
                             vm.NewTask.Status == K_STATUS_3 ? 3 : 0;
          
                int idtmp = Convert.ToInt32(vm.NewTask.ID);
                var task = new TaskModel_DB
                {
                    ID = idtmp,
                    TaskName = vm.NewTask.TaskName,
                    Status = 3,              // 假设 1=进行中，0=未开始
                    StartDate = DateTime.Parse(vm.NewTask.StartDate),
                    Consumed = 0,              // 已消耗时间（分钟？）
                    DataStatus = false,          // 数据已就绪
                    FilesStatus = false,         // 文件已下载
                    CheckStatus = false         // 尚未审核
                };
                string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
                var repo = new TaskRepository(connStr);
                var index = await repo.InsertCurrentTaskAsync(task);
                if (index > 0)
                {
                    var uiItem= DBTaskToTaskModel(task);
                    uiItem.TaskFinish += OnTaskFinshed;
                    uiItem.PopupShow += OnPopupShow;
                    uiItem.Ok += OnOk;
                    vm.CurrentTasks.Add(uiItem);
                
                }
                var sqlServer = new TaskRepository();
                var opModel = new OperationModel
                {
                    TaskID = idtmp,
                    TaskName = vm.NewTask.TaskName,
                    TimeStamp = DateTime.Now,
                };
                await sqlServer.InsertOperationAsync(opModel);

                var logModel = new LogModel
                {
                    TimeStamp =DateTime.Now,
                    UserName = vm.NewTask.TaskName,
                    Level = LogLevels.Info,
                    TaskId = vm.NewTask.ID,
                    TaskName = vm.NewTask.TaskName,
                    Message = $"The task {vm.NewTask.TaskName} has been generate by {vm.NewTask.TaskName}."

                };

                await sqlServer.InsertLogAsync(logModel);

            }


        }


        private TaskModel DBTaskToTaskModel(TaskModel_DB task)
        {
            var resoult = new TaskModel()
            {
                ID = task.ID.ToString(),
                TaskName = task.TaskName,
                Status = task.Status == 1 ? K_STATUS_1 :
                             task.Status == 2 ? K_STATUS_2 :
                             task.Status == 3 ? K_STATUS_3 : "Unknown",
                StartDate = task.StartDate.ToString("yyyy.MM.dd hh:mm:ss"),
                Timeing = task.Consumed.ToString(),
                DataStatus = task.DataStatus.ToString(),
                FilesStatus = task.FilesStatus.ToString(),
                CheckStatus = task.CheckStatus.ToString()


            };
            
            return resoult;
        }

        private void ClosePopup(object sender, RoutedEventArgs e)
        {
            vm.PopupIsOpen = false;
        }
   
        public void SetCurrentUser(User user)
        {
            _currentUser = user;
        }

        private async System.Threading.Tasks.Task LogAndHandleErrorAsync(TaskModel task, string message)
        {
            // ?? 建议：同样注入 TaskRepository 实例而不是在这里创建
            var sqlServer = new TaskRepository();
            var logModel = new LogModel
            {
                TaskId = task.ID,
                UserName = Global.User.UserName,
                TimeStamp = DateTime.Now,
                TaskName = task.TaskName,
                Level = LogLevels.Info, // 任务完成应该用 Info 级别，而不是 Error
                Message = message
            };

            // 如果插入日志失败，通知用户
            if (!await sqlServer.InsertLogAsync(logModel))
            {
                // 建议使用异步通知机制或更高级的日志框架
                System.Windows.MessageBox.Show("Add log failed. Please check the database connection.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private async System.Threading.Tasks.Task HandleDataProviderFinish(TaskRepository repo, TaskModel task, TaskModel_DB taskItem)
        {
            taskItem.DataStatus = true;

            if (await repo.UpdateCurrentTaskAsync(taskItem))
            {
                TaskModel taskToUpdate = vm.CurrentTasks.FirstOrDefault(t => t.ID == task.ID);
                if (taskToUpdate != null)
                {
                    // ?? 检查：如果 DataStatus 是 bool 类型，这里不应该是字符串 "True"
                    taskToUpdate.DataStatus = "True";
                }
            }

            // 调用提取的日志方法
            string logMsg = $"The data status of the {task.TaskName} is finished by DataProvider ({Global.User.UserName}).";
            await LogAndHandleErrorAsync(task, logMsg);
        }

        private async System.Threading.Tasks.Task HandlePptMakerFinish(TaskRepository repo, TaskModel task, TaskModel_DB taskItem)
        {
            if (taskItem.DataStatus)
            {
                taskItem.FilesStatus = true;

                if (await repo.UpdateCurrentTaskAsync(taskItem))
                {
                    TaskModel taskToUpdate = vm.CurrentTasks.FirstOrDefault(t => t.ID == task.ID);
                    if (taskToUpdate != null)
                    {
                        // ?? 检查：如果 FilesStatus 是 bool 类型，这里不应该是字符串 " True"
                        taskToUpdate.FilesStatus = " True";
                    }
                }

                string logMsg = $"The file status of the {task.TaskName} is finished by PptMaker ({Global.User.UserName}).";
                await LogAndHandleErrorAsync(task, logMsg);
            }
            else
            {
                System.Windows.MessageBox.Show("Data not ready, please check the data status first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async System.Threading.Tasks.Task HandleReviewerFinish(TaskRepository repo, TaskModel task, TaskModel_DB  taskItem)
        {
            if (taskItem.DataStatus && taskItem.FilesStatus)
            {
                taskItem.CheckStatus = true;
                taskItem.Status = 1;

                // 事务处理：InsertFinishedTaskAsync 和 DeleteCurrentTaskAsync 应该封装在一个数据库事务中
                // 以确保两个操作要么都成功，要么都失败。
                if (await repo.InsertFinishedTaskAsync(taskItem) && await repo.DeleteCurrentTaskAsync(taskItem))
                {
                    TaskModel taskToUpdate = vm.CurrentTasks.FirstOrDefault(t => t.ID == task.ID);
                    if (taskToUpdate != null)
                    {
                        taskToUpdate.FilesStatus = "True";
                        // taskToUpdate.CheckStatus = "True"; // 原代码注释掉了，但逻辑上应该更新
                        taskToUpdate.Status = K_STATUS_1;
                    }
                    // 更新 UI 列表
                    vm.CurrentTasks.Remove(task);
                    vm.FinishedTasks.Add(task);
                }

                string logMsg = $"The check status of the {task.TaskName} is finished and task moved to finished list by Reviewer ({Global.User.UserName}).";
                await LogAndHandleErrorAsync(task, logMsg);
            }
            else
            {
                System.Windows.MessageBox.Show("Data and file not ready, please check the data status first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Btn_Test_Click(object sender, RoutedEventArgs e)
        {
            TestEvent?.Invoke(this, EventArgs.Empty);
                
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
        public ObservableCollection<TaskModel> CurrentTasks { get; set; } = new ObservableCollection<TaskModel>();

        public ObservableCollection<TaskModel> FinishedTasks { get; set; } = new ObservableCollection<TaskModel>();

        private TaskModel _newTask = new TaskModel();
        public TaskModel NewTask { get => _newTask; set { _newTask = value; RaisePropertyChanged(); } }
    }

    public class TaskModel : ObeservableObject
    {
        public event EventHandler<TaskModel> TaskFinish;
        public event EventHandler PopupShow;
        public event EventHandler<TaskModel> Ok;

        private string _id;
        public string ID
        {
            get { return _id; }
            set { _id = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

        public string _taskName;
       public string TaskName
        {
            get { return _taskName; }
            set { _taskName = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }


        private string _status;
        public string Status
        {
            get { return _status; }
            set { _status = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }
        
     
        public string _startDate;
        public string StartDate
        {
            get { return _startDate; }
            set { _startDate = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }


        public string Timeing{ get; set; }
   
        public string _dataStatus;
        public string DataStatus
        {
            get { return _dataStatus; }
            set { _dataStatus = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

        public string _filesStatus;
        public string FilesStatus
        {
            get { return _filesStatus; }
                
            set { _filesStatus = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

        public string _checkStatus;
        public string CheckStatus
        {
            get { return _checkStatus; }
            set { _checkStatus = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

        //public string DataStatus { get; set; }
        //public string FilesStatus { get; set; }
        //public string CheckStatus { get; set; }

        public ICommand FinishCommand { get; }

        public ICommand PopupShowCommand { get; }
        public ICommand OkCommand { get; }
        public TaskModel(string id="",string taskName="")
        {

            ID = id;
            TaskName = taskName;
            Status = "Not Started";

            StartDate = DateTime.Now.ToString("yyyy.MM.dd");
            Timeing = "0 ";


            DataStatus = "Not Started";
            FilesStatus = "Not Started";
            CheckStatus = "Not Started";
            FinishCommand = new RelayCommand<TaskModel>(OnFinishTask);
            PopupShowCommand = new RelayCommand(OnShowPopup);
            OkCommand = new RelayCommand<TaskModel>(OnOk);
        }
        [RelayCommand]
        private void OnFinishTask(TaskModel task)
        {
            if (task != null)
            {
                // 处理完成逻辑
                var result = System.Windows.MessageBox.Show(
                                $"Are you sure to finish the current step of {task.TaskName}?",
                                "Confirm Finish", // 建议加上标题
                                MessageBoxButton.OKCancel, // ? 用 OKCancel，不是 OK
                                MessageBoxImage.Question
                            );

                // ? 只有用户点击 OK 时才继续
                if (result == MessageBoxResult.OK)
                {
                    // 更新状态
                    //task.Status = "Finished";

                    // 触发事件
                    TaskFinish?.Invoke(this, task);
                }

                
            }
        }


        [RelayCommand]

        private void OnShowPopup()
        {
            PopupShow?.Invoke(this, EventArgs.Empty);

        }

        [RelayCommand]

        private void OnOk(TaskModel task)
        {
            //if(Status == "Not Started")
            //  {
            //    Status = "In Progress";
            //  }
                
            Ok?.Invoke(this, task);

        }

    }


}
