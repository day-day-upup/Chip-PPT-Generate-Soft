using ScottPlot.WPF;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ChipManualGenerationSogt
{
    /// <summary>
    /// App.xaml �Ľ����߼�
    /// </summary>
    public partial class App : Application
    {

        // 1. ��ʾ��¼����

        protected override async void OnStartup(StartupEventArgs e)
        {
            this.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            base.OnStartup(e);

            // 1. ��ʾ��¼����
            //var loginWindow = new LoginW();
            //loginWindow.Show(); // ģ̬��ʾ
            //var mainWindow = new MainWindow();
            //mainWindow.Show();
            //string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
            //var task = new TaskModel2
            //{
            //    ID = 1001,
            //    TaskName = "ͼ��������",
            //    Status = 1,                 // ���� 1=�����У�0=δ��ʼ
            //    StartDate = DateTime.Now,
            //    Consumed = 30,              // ������ʱ�䣨���ӣ���
            //    DataStatus = true,          // �����Ѿ���
            //    FilesStatus = true,         // �ļ�������
            //    CheckStatus = false         // ��δ���
            //};
            //var repo = new TestRecordRepository(connStr);
            //repo.SaveTask(connStr, task);

            //string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";

            //var userRepository = new UserRepository(connStr);
            //var tasRepository = new TaskRepository(connStr);
            //// 3. ���÷�����ע�⣺���� async ��������
            //try
            //{
            //    //List<User> users = await userRepository.GetAllUsersAsync();

            //    //foreach (var user in users)
            //    //{
            //    //    //Console.WriteLine($"{user.ID}: {user.Name}, Priority: {user.Priority}");
            //    //}
            //    List<TaskModel_DB> users = await tasRepository.GetAllFinishedTasksAsync();

            //    foreach (var user in users)
            //    {
            //        //Console.WriteLine($"{user.ID}: {user.Name}, Priority: {user.Priority}");
            //    }
            //    users = await tasRepository.GetAllCurrentTasksAsync();

            //    foreach (var user in users)
            //    {
            //        //Console.WriteLine($"{user.ID}: {user.Name}, Priority: {user.Priority}");
            //    }

            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"�����û�ʧ��: {ex.Message}");
            //}

            //var control = new WpfPlot();
            //control.Show();

            var user = new User
            {
                ID = 1001,
                UserName = "admin",
                Password = "123456",
                
            };
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string dir = System.IO.Path.Combine(appDir, "resources", "settings");
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }
            string filePath = System.IO.Path.Combine(appDir, "resources", "settings", "settings.json");
            if (!System.IO.File.Exists(filePath))
            {

                var login = new LoginW();
                if (login.ShowDialog() == true)
                {
                    var settings = new Settings()
                    {
                        LastLoggedInUser = login.User.UserName,
                        LastLoggedInPassword = login.User.Password,

                    };

                    await JsonPaser.CreateSettingsJsonFile(filePath, settings);
                    Global.User = login.User;
                    var main = new MainWindow(Global.User);

                    main.SetCurrentUserTaskManage(Global.User);
                    main.SetUsersNameLogWin(login.AllUser);
                    main.Show();
                    // ��������ʾ�󣬻ָ�Ĭ���˳���Ϊ����ѡ��
                    this.ShutdownMode = ShutdownMode.OnLastWindowClose;


                }
                else
                {
                    //Shutdown(); // �û�ȡ����¼
                }


            }
            else
            {
                var settings = await JsonPaser.ReadSessionJsonAsync(filePath);
                var login = new LoginW(settings.LastLoggedInUser, settings.LastLoggedInPassword);
                if (login.ShowDialog() == true)
                {
                    if (login.User.UserName != settings.LastLoggedInUser)
                    {
                        var newSettings = new Settings()
                        {
                            LastLoggedInUser = login.User.UserName,
                            LastLoggedInPassword = login.User.Password,

                        };

                        await JsonPaser.CreateSettingsJsonFile(filePath, newSettings);

                    }
                    Global.User = login.User;
                    var main = new MainWindow(Global.User);
                    main.SetCurrentUserTaskManage(Global.User);
                    main.SetUsersNameLogWin(login.AllUser);
                    main.Show();
                    // ��������ʾ�󣬻ָ�Ĭ���˳���Ϊ����ѡ��
                    this.ShutdownMode = ShutdownMode.OnLastWindowClose;


                }
                else
                {
                    //Shutdown(); // �û�ȡ����¼
                }

            }


            //var main = new MainWindow(user);
            //main.Show();

            //var login = new LoginW();
            //if (login.ShowDialog() == true)
            //{
            //    var main = new MainWindow(login.User);
            //    main.Show();
            //    // ��������ʾ�󣬻ָ�Ĭ���˳���Ϊ����ѡ��
            //    this.ShutdownMode = ShutdownMode.OnLastWindowClose;
            //}
            //else
            //{
            //    Shutdown(); // �û�ȡ����¼
            //}

        }

    }






    
}
