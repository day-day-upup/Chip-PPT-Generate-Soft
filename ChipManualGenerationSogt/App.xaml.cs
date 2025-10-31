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
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {

        // 1. 显示登录窗口

        protected override async void OnStartup(StartupEventArgs e)
        {
            this.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            base.OnStartup(e);

            // 1. 显示登录窗口
            //var loginWindow = new LoginW();
            //loginWindow.Show(); // 模态显示
            //var mainWindow = new MainWindow();
            //mainWindow.Show();
            //string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
            //var task = new TaskModel2
            //{
            //    ID = 1001,
            //    TaskName = "图像处理任务",
            //    Status = 1,                 // 假设 1=进行中，0=未开始
            //    StartDate = DateTime.Now,
            //    Consumed = 30,              // 已消耗时间（分钟？）
            //    DataStatus = true,          // 数据已就绪
            //    FilesStatus = true,         // 文件已下载
            //    CheckStatus = false         // 尚未审核
            //};
            //var repo = new TestRecordRepository(connStr);
            //repo.SaveTask(connStr, task);

            //string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";

            //var userRepository = new UserRepository(connStr);
            //var tasRepository = new TaskRepository(connStr);
            //// 3. 调用方法（注意：这是 async 方法！）
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
            //    Console.WriteLine($"加载用户失败: {ex.Message}");
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
                    // 主窗口显示后，恢复默认退出行为（可选）
                    this.ShutdownMode = ShutdownMode.OnLastWindowClose;


                }
                else
                {
                    //Shutdown(); // 用户取消登录
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
                    // 主窗口显示后，恢复默认退出行为（可选）
                    this.ShutdownMode = ShutdownMode.OnLastWindowClose;


                }
                else
                {
                    //Shutdown(); // 用户取消登录
                }

            }


            //var main = new MainWindow(user);
            //main.Show();

            //var login = new LoginW();
            //if (login.ShowDialog() == true)
            //{
            //    var main = new MainWindow(login.User);
            //    main.Show();
            //    // 主窗口显示后，恢复默认退出行为（可选）
            //    this.ShutdownMode = ShutdownMode.OnLastWindowClose;
            //}
            //else
            //{
            //    Shutdown(); // 用户取消登录
            //}

        }

    }






    
}
