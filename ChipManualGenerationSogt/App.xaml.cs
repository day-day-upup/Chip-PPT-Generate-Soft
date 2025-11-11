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
                    Shutdown(); // 用户取消登录
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
                    //Application.Current.Shutdown();
                    Shutdown(); // 用户取消登录
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
