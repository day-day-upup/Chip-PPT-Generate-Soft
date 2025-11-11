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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ChipManualGenerationSogt
{
    /// <summary>
    /// LeftSider.xaml µÄ½»»¥Âß¼­
    /// </summary>
    public partial class LeftSider : UserControl
    {
        LeftSierModel vm;
        public event EventHandler OnHomeBtnClick;
        public event EventHandler OnOperationBtnClick;
        public event EventHandler OnLogBtnClick;
        public event EventHandler OnAddBtnClick;
        public LeftSider()
        {

            InitializeComponent();
            vm = new LeftSierModel();
            this.DataContext = vm;
            //vm.UserName =;
        }
        public void SetUserName(string name)
        {
            vm.UserName = name;
        }

        private void Btn_Home_Clicked(object sender, RoutedEventArgs e)
        {
            OnHomeBtnClick?.Invoke(this, EventArgs.Empty);
        }

        private void Btn_Operation_Clicked(object sender, RoutedEventArgs e)
        {
            OnOperationBtnClick?.Invoke(this, EventArgs.Empty);
        }

        private void Btn_Log_Clicked(object sender, RoutedEventArgs e)
        {
            OnLogBtnClick?.Invoke(this, EventArgs.Empty);
        }

        private void Btn_Add_Clicked(object sender, RoutedEventArgs e)
        {
            OnAddBtnClick?.Invoke(this, EventArgs.Empty);
        }
    }
    class LeftSierModel : ObeservableObject
    {

        string userName;
        public string UserName
        {
            get { return userName; }
            set { userName = value; RaisePropertyChanged(); }
        }

        public string Version { get; set; } = Properties.Settings.Default.Version;
    }
}
