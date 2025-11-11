using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace ChipManualGenerationSogt
{
    /// <summary>
    /// LoginW.xaml µÄ½»»¥Âß¼­
    /// </summary>
    public partial class LoginW : Window
    {
        LogInModel vm ;
        public User User { get; set; }
        public List<string> AllUser { get;set; } =  new List<string>();
        public LoginW()
        {
            InitializeComponent();
            vm = new LogInModel();
            this.DataContext = vm;
        }

        public LoginW( string userName, string passworld)
        {
            InitializeComponent();
            vm = new LogInModel();
            this.DataContext = vm;
            vm.UserName = userName;
            txtPassword.Password = passworld;
        }


        private async void Btn_LogIn_Clicked(object sender, RoutedEventArgs e)
        {
            string connStr = "Server=192.168.1.209;Database=mlChips;User ID=sa;Password=qotana;Encrypt=false;TrustServerCertificate=true;";
            vm.IsBusy = true;
            vm.BusyMessage = "Logining...";
            var repo = new TestRecordRepository(connStr);
            var users = repo.GetUsers(connStr);
            foreach (var user in users)
            {
                AllUser.Add(user.UserName);
            }
            foreach (var user in users)
            {

                if (user.UserName == vm.UserName && user.Password == txtPassword.Password)
                {
                    Console.WriteLine("Login Success");
                    var sqlServer = new TaskRepository();
                    var logmodel = new LogModel 
                    {
                        UserName = user.UserName,
                        Message = "Login Success",
                        TimeStamp = DateTime.Now,
                        Level = LogLevels.Info
                    };

                     await sqlServer.InsertLogAsync(logmodel);
                    User = user;
                    this.DialogResult = true;
                    
                    break;
                }
                else
                {
                    

                }
                
            }
            if (this.DialogResult != true)
            {
                MessageBox.Show("User Name or Password is incorrect", "", MessageBoxButton.OK, MessageBoxImage.Error);
                //var sqlServer = new TaskRepository();
                //var logmodel = new LogModel
                //{
                //    UserName = user.UserName,
                //    Message = "Login Success",
                //    TimeStamp = DateTime.Now,
                //    Level = LogLevels.Info
                //};

                //await sqlServer.InsertLogAsync(logmodel);
            }
            vm.IsBusy = false;
        }
    }

    public class LogInModel : ObeservableObject
    { 
        private string _userName;
        public String UserName
        {
            get { return _userName; }
            set { _userName = value; RaisePropertyChanged(); }
        }

        private string _passworld;

        public string Passworld
        { 
             get { return _passworld; }
             set { _passworld = value; RaisePropertyChanged(); }
        }

        public DateTime _startTime;
        public DateTime StartTime
        {
            get { return _startTime; }
            set { _startTime = value; RaisePropertyChanged(); }
        }

        public DateTime _endTime;
        public DateTime EndTime
        {
            get { return _endTime; }
            set { _endTime = value; RaisePropertyChanged(); }
        }


        bool _isBusy;
        public bool IsBusy
        {
            get { return _isBusy; }
            set { _isBusy = value; RaisePropertyChanged(nameof(IsBusy)); }
        }

        private string _busyMessage;
        public string BusyMessage
        {
            get { return _busyMessage; }
            set { _busyMessage = value; RaisePropertyChanged(nameof(BusyMessage)); }
        }



    }

}
