using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
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

namespace ChipManualGenerationSogt
{
    /// <summary>
    /// LogWin.xaml 的交互逻辑
    /// </summary>
    public partial class LogWin : UserControl
    {
        LogWinModel vm;
        public event EventHandler BackEvent;
        public LogWin()
        {
            InitializeComponent();

            vm = new LogWinModel();
            DataContext = vm;
            vm.LogText = "这是一个日志测试:A Log in\n 这是一个日志测试:A Select Amplifier MM809\n  这是一个日志测试:A Enter SN:L004x,ON:L004x\n 这是一个日志测试： A Select contion:VG 4.0V,VD 1.5V\n 这是一个日志测试： A Select Filter:MML806\n 这是一个日志测试： A Select Amplifier MM809\n 这是一个日志测试： A Select contion:VG 4.0V,VD 1.5V\n 这是一个日志测试： A Select Filter:MML806\n 这是一个日志测试： A Select Amplifier MM809\n 这是一个日志测试： A Select contion:VG 4.0V,VD 1.5V\n 这是一个日志测试： A Select Filter:MML806\n 这是一个日志测试： A Select Amplifier MM809\n 这是一个日志测试： A Select contion:VG 4.0V,Idd:67mA\n 这是一个日志测试： A Select Filter:M";
            vm.StartTime = DateTime.Now.AddDays(-1);
            vm.EndTime = DateTime.Now;
            vm.LevelList.Add("DEBUG");
            vm.LevelList.Add("INFO");
            vm.LevelList.Add("WARNING");
            vm.LevelList.Add("ERROR");
        }

        private async void Btn_Query_Clicked(object sender, RoutedEventArgs e)
        {
            try
            {
                var sqlServer = new TaskRepository();
                var result = await sqlServer.QueryLogsAsync(vm.SelectedUser,
                    vm.TaskId, vm.Level,null, null, null,vm.StartTime,vm.EndTime);
                vm.LogText = LogModelToText(result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void SetComboxSourse(List<string> source)
        {
            vm.UserList.Clear();
            foreach (var item in source)
            {
                vm.UserList.Add(item);
            }
        }

        private string LogModelToText( List<LogModel> models)
        {
            var resoult = "";
            foreach (var model in models)
            {
                resoult += "TimeStamp: " + model.TimeStamp.ToString() + ";  ";
                resoult += "Task ID: " + model.TaskId + ";  ";
                resoult += "Task Name: " + model.TaskName + ";  ";
                resoult += "Level：" + model.Level + ";  ";
                resoult += "Message：" + model.Message + "\n";
            }
            return resoult;
        
        }

        private void Btn_Back_Click(object sender, RoutedEventArgs e)
        {
            BackEvent?.Invoke(this, EventArgs.Empty);
        }
    }

    public class LogWinModel : ObeservableObject
    {
        private string _logText;
        public String LogText
        {
            get { return _logText; }
            set { _logText = value; RaisePropertyChanged(); }
        }
        private string _selectedUser;

        public string SelectedUser
        {
            get { return _selectedUser; }
            set { _selectedUser = value; RaisePropertyChanged(); }
        }

        public ObservableCollection<string> UserList { get; set; } = new ObservableCollection<string>();


        private string _taskId;
        public string TaskId
        {
            get { return _taskId; }
            set { _taskId = value; RaisePropertyChanged(); }
        }


        private string _taskName;
        public string TaskName
        {
            get { return _taskName; }
            set
            {
                _taskName = value; RaisePropertyChanged();
            }
        }

        private string _level;
        public string Level
        {
            get { return _level; }
            set { _level = value; RaisePropertyChanged(); }
        }
        public ObservableCollection<string> LevelList { get; set; } = new ObservableCollection<string>();

        private DateTime _startTime;
        public DateTime StartTime
        {
            get { return _startTime; }
            set { _startTime = value; RaisePropertyChanged(); }
        }

        private DateTime _endTime;
        public DateTime EndTime
        {
            get { return _endTime; }
            set { _endTime = value; RaisePropertyChanged(); }
        }

    }


}
