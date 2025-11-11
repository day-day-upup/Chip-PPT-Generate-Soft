using System;
using System.Collections.Generic;
using System.IO;
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
    /// PdfShowWin.xaml 的交互逻辑
    /// </summary>
    public partial class PdfShowWin : Window
    {
        static public string PdfPath { set; get; }
        static public string PPTPath { set; get; }
        public bool Status { set; get; }//true-- 表示是预览状态， false--表示是正式生成状态
        public PdfShowWin()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;
        }

        public void ShowPdf(string pdfPath)
        {
            PdfPath= pdfPath;
            if (!File.Exists(pdfPath))
            {
                
                throw new FileNotFoundException($"File not found: {pdfPath}");
                
            }
            //webView.Source = new Uri($"file:///{pdfPath.Replace("\\", "/")}");
            webView.Source = new Uri(@"file:///F:\PROJECT\ChipManualGeneration\exe\ChipManualGenerationSogt\ChipManualGenerationSogt\bin\Debug\resources\files\demo.pdf");
        }


        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 你的自定义关闭逻辑
            //var result = MessageBox.Show(
            //    "确定要关闭窗口吗？",
            //    "确认",
            //    MessageBoxButton.OKCancel,
            //    MessageBoxImage.Question
            //);

            //if (result == MessageBoxResult.Cancel)
            //{
            //    e.Cancel = true; // 取消关闭
            //}
            if (Status)
            {
                //预览状态， 关闭时，删除临时文件
                try
                {
                    File.Delete(PdfPath);
                    File.Delete(PPTPath);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
                
            }
                
            
        }
    }


    
}
