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
        public event EventHandler    AddEvent;
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
                PPTModel = item.PPTModel,
                TaskName = item.TaskName,
                Status = item.Status,
                Level = item.Level.Trim(),
                DataStatus = item.DataStatus,
                FilesStatus = item.FilesStatus,
                StartDate = item.StartDate,
                EndDate = item.EndDate,
                Conditions = item.Conditions,
                Major = item.Major,
                Minor = item.Minor,
            };
            return uiItem;
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
        private void OnDetaileEvent(object sender, TaskTableItem task)
        {
            vm.PopupIsOpen = true;
            vm.PopupTitle = task.TaskName;
            
            vm.PopupTopMessage = task.Conditions;
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
            var task = selectedItem as TaskTableItem; // 替换为你的实际类型
            //MessageBox.Show($"双击了任务：{task.TaskName}");
            Console.WriteLine($"{task.TaskName}");
            Global.TaskModel = task;
            
            
            //TaskExcute?.Invoke(this, task);
            
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
                //Timeing = task.Consumed.ToString(),
                DataStatus = task.DataStatus.ToString(),
                FilesStatus = task.FilesStatus.ToString(),
                //CheckStatus = task.CheckStatus.ToString()
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

        private async void Btn_Test_Click(object sender, RoutedEventArgs e)
        {


            string ftpRootPath = "Amplifier/MML806";

            //string folderPath = "F:\\PROJECT\\ChipManualGeneration\\原始数据\\MML004X 手册数据-3";
            //await FtpClient.UploadFolderAsync(folderPath, ftpRootPath);

            await FtpClient.DownloadFolderAsync(ftpRootPath, Global.FileBasePath);
           

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
                //uiItem.TaskFinish += OnTaskFinshed;
                //uiItem.PopupShow += OnPopupShow;
                //uiItem.Ok += OnOk;
                uiItem.DetaileEvent += OnDetaileEvent;
                //vm.CurrentTasks.Add(uiItem);
                vm.TableItemSources.Add(uiItem);
            }
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

        private ObservableCollection<object> _selectedDeviceMajors = new ObservableCollection<object>();
        public ObservableCollection<object> SelectedDeviceMajors
        {
            get { return _selectedDeviceMajors; }
            set
            {
                _selectedDeviceMajors = value;
                // 确保你的 RaisePropertyChanged 方法存在
                RaisePropertyChanged();
            }
        }
        public ObservableCollection<string> DeviceMajorSoureces { get; set; } = new ObservableCollection<string>() ;



        private ObservableCollection<object> _selectedDeviceMinors = new ObservableCollection<object>();
        public ObservableCollection<object> SelectedDeviceMinors
        {
            get { return _selectedDeviceMinors; }
            set
            {
                _selectedDeviceMinors = value;
                // 确保你的 RaisePropertyChanged 方法存在
                RaisePropertyChanged();
            }
        }
        public ObservableCollection<string> DeviceMinorSoureces { get; set; } = new ObservableCollection<string>() ;


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

        private string _chipNumber;
        public string ChipNumber
        {
            get { return _chipNumber; }
            set { _chipNumber = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
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

        private string _level;
        public string Level
        {
            get { return _status; }
            set { _status = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

        private string _major;

        public string Major
        {
            get { return _major; } 
            set{_major = value; RaisePropertyChanged(); }
 
        }


        private string _minor;

        public string Minor
        {
            get { return _minor; }
            set { _minor = value; RaisePropertyChanged(); }

        }

        public string _startDate;
        public string StartDate
        {
            get { return _startDate; }
            set { _startDate = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

        public string _endDate;
        public string EndDate
        {
            get { return _endDate; }
            set { _endDate = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

   
        private string _dataStatus;
        public string DataStatus        
        {
            get { return _dataStatus; }
            set { _dataStatus = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }


        private string _filesStatus;
        public string FilesStatus
        {
            get { return _filesStatus; }
                
            set { _filesStatus = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！
        }

        private string _condition;

        public  string Condition
        {
           get { return _condition; }
            set { _condition = value; RaisePropertyChanged(); } // ← 没有 RaisePropertyChanged()！

        }


        public ICommand FinishCommand { get; }

        public ICommand PopupShowCommand { get; }
        public ICommand OkCommand { get; }
        public TaskModel(string id="",string taskName="")
        {

            ID = id;
            TaskName = taskName;
            Status = "Not Started";

            StartDate = DateTime.Now.ToString("yyyy.MM.dd");
            EndDate = "NULL";

            DataStatus = "Not Started";
            FilesStatus = "Not Started";

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

        }

        // --- 封装所有需要通知 UI 的属性 ---


        public int ID
        {
            get => DataModel.ID;
           
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

        //public TaskTableItem()
        //{
        //}
    }
}
