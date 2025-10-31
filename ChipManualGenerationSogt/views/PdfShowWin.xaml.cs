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
    /// PdfShowWin.xaml �Ľ����߼�
    /// </summary>
    public partial class PdfShowWin : Window
    {
        static public string PdfPath { set; get; }
        static public string PPTPath { set; get; }
        public bool Status { set; get; }//true-- ��ʾ��Ԥ��״̬�� false--��ʾ����ʽ����״̬
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
            webView.Source = new Uri($"file:///{pdfPath.Replace("\\", "/")}");
            //webView.Source = new Uri(@"file:///F:\PROJECT\ChipManualGeneration\exe\ChipManualGenerationSogt\ChipManualGenerationSogt\bin\Debug\resources\files\demo.pdf");
        }


        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // ����Զ���ر��߼�
            //var result = MessageBox.Show(
            //    "ȷ��Ҫ�رմ�����",
            //    "ȷ��",
            //    MessageBoxButton.OKCancel,
            //    MessageBoxImage.Question
            //);

            //if (result == MessageBoxResult.Cancel)
            //{
            //    e.Cancel = true; // ȡ���ر�
            //}
            if (Status)
            {
                //Ԥ��״̬�� �ر�ʱ��ɾ����ʱ�ļ�
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
