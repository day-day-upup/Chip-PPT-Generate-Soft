using DocumentFormat.OpenXml.Packaging;
using ScottPlot.Finance;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using static ChipManualGenerationSogt.Curves;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Threading;
using Microsoft.Win32;
using ChipManualGenerationSogt.models;
using Microsoft.Web.WebView2.Core;
using System.Runtime.InteropServices;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.Vml;
using ScottPlot;
using Microsoft.Office.Interop.PowerPoint;
using ScottPlot.AxisLimitManagers;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Security.Principal;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Bibliography;
namespace ChipManualGenerationSogt
{
    /// <summary>
    /// MainWindow.xaml �Ľ����߼�
    /// </summary>
    public partial class MainWindow : Window
    {
        MainWindowModel vm;
        PptDataModel pptDataModel;
        TaskModel _task;
        List<PlotModel> plots;
        User _user;
        List<string> _plotNames = new List<string>();
        public MainWindow(User user)
        {
            InitializeComponent();
            vm = new MainWindowModel();
            this.DataContext = vm;
            pptDataModel = new PptDataModel();
            vm.PopupVisible = false;
            filter.OnQueryDatabaseFinished += HandleQueryFinished;
            leftSider.SetUserName(user.UserName);
            leftSider.OnHomeBtnClick += HandleHomeBtnClick;
            leftSider.OnLogBtnClick += HandleLogBtnClick;
            leftSider.OnOperationBtnClick += HandleOperationBtnClick;
            leftSider.OnAddBtnClick += HandleAddBtnClick;

            plots = new List<PlotModel>();
            //vm.LogText = "����һ����־����:A Log in\n ����һ����־����:A Select Amplifier MM809\n  ����һ����־����:A Enter SN:L004x,ON:L004x\n ����һ����־���ԣ� A Select contion:VG 4.0V,VD 1.5V\n ����һ����־���ԣ� A Select Filter:MML806\n ����һ����־���ԣ� A Select Amplifier MM809\n ����һ����־���ԣ� A Select contion:VG 4.0V,VD 1.5V\n ����һ����־���ԣ� A Select Filter:MML806\n ����һ����־���ԣ� A Select Amplifier MM809\n ����һ����־���ԣ� A Select contion:VG 4.0V,VD 1.5V\n ����һ����־���ԣ� A Select Filter:MML806\n ����һ����־���ԣ� A Select Amplifier MM809\n ����һ����־���ԣ� A Select contion:VG 4.0V,Idd:67mA\n ����һ����־���ԣ� A Select Filter:M";
            taskMangeControls.TaskExcute += HandleTaskExcute;
            taskMangeControls.TestEvent += async(sender, e) =>
            {
                try
                {
                    string appDir = AppDomain.CurrentDomain.BaseDirectory;
                    string picsPath = System.IO.Path.Combine(appDir, "resources", "pic");
                    System.IO.Directory.Delete(picsPath, true);
                    System.IO.Directory.CreateDirectory(picsPath);
                    curves.SaveAllPlot(picsPath);
                    PptDataModeFactory();


                    string pptFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pptx");
                    string pdfFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pdf");
                    await Task.Run(() => GeneratePPT(pptFile));
                    await Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));
                    var pdfShown = new PdfShowWin();
                    //pdfShown.Status = true;
                    //PdfShowWin.PPTPath = pptFile;
                    //PdfShowWin.PdfPath = pdfFile;
                    pdfShown.ShowPdf(pdfFile);
                    pdfShown.Show();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            };
            _task = new TaskModel();

            //var win = new LoginW();
            //win.Show();

            //vm.Records.Add(new NewTreeViewItem
            //{
            //    Content = "Amplifier",
            //    Visible = true,
            //    // Children.Add(new NewTreeViewItem { Content = "�����ձ�" }),
            //    Children =
            //        {
            //            new NewTreeViewItem{ Content = "MML806" },
                        
            //        }
            //});
            //vm.Records.Add(new NewTreeViewItem
            //{
            //    Content = "Filter",
            //    Visible = true,
            //    Children =
            //        {
            //            new NewTreeViewItem{ Content = "MML806" },
            //            new NewTreeViewItem { Content = "MML807" }
            //        }
            //});

            //var AmplifierCollection = Properties.Settings.Default.Amplifier;

            //foreach (var item in AmplifierCollection)
            //{
            //    Console.WriteLine(item);
            //}
            LoadProperites();
            //Directory.Delete(@"CopiedReports", true);
            //QueryDatabase("L004X", "L004X", "", "");
            ////test();
            //test2();

            //var pdfShown = new PdfShowWin();
            //pdfShown.ShowPdf("C:\\Users\\pengyang\\Desktop\\books\\UNIX�������.pdf");
            //pdfShown.Show();

            //test3();
            //PPTChange();
        }
        public void PptDataModeFactory()
        {
            var imgs = images.GetAllImage();
            //********************ĸ����Ϣ
            var tablesModel = table.GetModel();
            var basicTable = table.GetBasicTableData();
            if (basicTable.Count == 5)
            {

                pptDataModel.SliderMaster = new SliderMasterModel
                {
                    TopPN = basicTable.ElementAt(0).info,
                    Version = basicTable.ElementAt(1).info,
                    ProductName = basicTable.ElementAt(2).info,
                    FrequencyRange = basicTable.ElementAt(3).info,
                    RPN = basicTable.ElementAt(0).info,
                    RightBarInfo = basicTable.ElementAt(4).info
                };
            }
            else
            {
                throw new Exception("Basic Table Fail Not Right");

            }
            #region ��һҳ
            string features = "Featrues\n";
            foreach (var item in table.GetFeatureTableData())
            {
                features += "\u2022" +"    " +item.name + " : " + item.info + "\n";
            }
            features = features.TrimEnd('\n');

            //�����ƥ�丳ֵ��ʽ���ܻ���Ҫ����

            string elecCondition = "TA = +25\u2103, " + filter.GetFileterCondition().VD_VG_Conditon.ElementAt(0).Replace('&', ',');

            pptDataModel.FirstPage = new FirstPageModel
            {
                FeaturesText = features,
                TypicalApplicationsText = "Typical Applications\n\u2022    Test Instrumentation\n\u2022    Microwave Radio & VSAT\n\u2022    Military & Space\n\u2022    Telecom Infrastructure\n\u2022    Fiber Optics",
                ElectricalSpecsTitle = "Electrical Specifications",
                ElectricalSpecsCondition = elecCondition,
                ParameterTableData = table.GetParameterTableInfo(),
                FunctionalBlockDiagramImage = new ImageModel
                {
                    ImagePath = imgs.ElementAt(0).filePath, 
                    ImageName = imgs.ElementAt(0).name,
                    Width = 2_500_000,
                    Height = 2_000_000,

                    XPoistion = 2_500_000,
                    YPoistion = 2_000_000,
                }
            };
            #endregion



            #region ����ҳ
            // ���ֻ��ַ�ʽ�ο���MML806_V3.pptx
            pptDataModel.CurvesImagePage = new CurvesImagePageModel();
            //string curveT1 = "Measurement Plots: S-parameters\n" + filter.GetFileterCondition().VD_VG_Conditon.ElementAt(0).Replace('&', ',');
            //pptDataModel.CurvesImagePage.CurveTitles.Add(curveT1);


            //string curveT25_1 = "Measurement Plots: S-parameters\n " + "TA = +25\u2103";
            //string curveT25_2 = "Measurement Plots: P1dB\n TA = +25\u2103";
            //string curveT25_3 = "Measurement Plots: OIP3\n TA = +25\u2103";
            //string curveT25_4 = "Measurement Plots: Psat\n TA = +25\u2103";
            //string curveT25_5 = "Measurement Plots: Noise Figure\n TA = +25\u2103";
            //pptDataModel.CurvesImagePage.CurveTitles.Add(curveT25_1);
            //pptDataModel.CurvesImagePage.CurveTitles.Add(curveT25_2);
            //pptDataModel.CurvesImagePage.CurveTitles.Add(curveT25_3);
            //pptDataModel.CurvesImagePage.CurveTitles.Add(curveT25_4);
            //pptDataModel.CurvesImagePage.CurveTitles.Add(curveT25_5);


            //List<string> curveE_1 = new List<string>();
            //List<string> curveE_2 = new List<string>();
            //List<string> curveE_3 = new List<string>();
            //List<string> curveE_4 = new List<string>();
            //List<string> curveE_5 = new List<string>();
            //foreach (var item in filter.GetFileterCondition().VD_VG_Conditon)
            //{
            //    string str1 = "Measurement Plots: S-parameters\n" + item.Replace('&', ',');
            //    string str2 = "Measurement Plots: P1dB\n" + item.Replace('&', ',');
            //    string str3 = "Measurement Plots: OIP3\n " + item.Replace('&', ',');
            //    string str4 = "Measurement Plots: Psat\n" + item.Replace('&', ',');
            //    string str5 = "Measurement Plots: Noise Figure\n" + item.Replace('&', ',');
            //    pptDataModel.CurvesImagePage.CurveTitles.Add(str1);
            //    pptDataModel.CurvesImagePage.CurveTitles.Add(str2);
            //    pptDataModel.CurvesImagePage.CurveTitles.Add(str3);
            //    pptDataModel.CurvesImagePage.CurveTitles.Add(str4);
            //    pptDataModel.CurvesImagePage.CurveTitles.Add(str5);

            //}
            pptDataModel.CurvesImagePage.CurveTitles = _plotNames;
            pptDataModel.CurvesImagePage.CurveImagesPath = curves.GetAllCurvesImagesFilePath();

            #endregion


            #region ��������ҳ

            pptDataModel.EndToFront5Page = new EndToFront5();
            pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTableTitle = "Absolute Maximum Ratings";
            pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTable = DataConverter.ConvertListToTwoDArray(table.GetAbsoluteRatingsData());
            pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTableTitle = "Typical Supply Current";
            

            pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTable = DataConverter.ConvertThreeElementListToTwoDArray(table.GetCurrentVdVgData());
            pptDataModel.EndToFront5Page.WarningImage = new ImageModel
            {
                ImagePath = @"F:\PROJECT\ChipManualGeneration\exe\2.png",
                Height = 50_0000,
                Width = 50_0000,
            };
            pptDataModel.EndToFront5Page.WarningText = "ELECTROSTATIC SENSITIVE DEVICE\n OBSERVE HANDLING PRECAUTIONS";
            #endregion

            #region  ������4ҳ

            pptDataModel.EndToFront4Page = new EndToFront4();
            pptDataModel.EndToFront4Page.PinImage = new ImageModel
            {
                ImagePath = imgs.ElementAt(1).filePath,
                ImageName = imgs.ElementAt(1).name,
                Height = 500_0000,
                Width = 500_0000
            };
            string notes = "";
            foreach (var item in table.GetNotesData())
            {
                notes += item +"\n";
            }

            pptDataModel.EndToFront4Page.NoteText = notes.TrimEnd('\n').Replace('?','\u00B2');

            #endregion

            #region ������3ҳ
            pptDataModel.EndToFront3Page = new EndToFront3();
            pptDataModel.EndToFront3Page.StructImage = new ImageModel
            {
                ImageName = imgs.ElementAt(2).name,
                ImagePath = imgs.ElementAt(2).filePath,
                Height = 550_0000,
                Width = 350_0000
            };
            List<(string,string)> d1=new List<(string,string)>();
            (string, string) hearderD1 = ("Item", "Description");
            d1.Add(hearderD1);
            foreach (var item in table.GetDescription1Data())
            {
                d1.Add(item);
            }
            List<(string, string, string)> d2 = new List<(string, string, string)>();
            (string, string, string)hearderD2 = ("Item", "Funciton", "Description");
            d2.Add(hearderD2);
            foreach (var item in table.GetDescription2Data())
            {
                d2.Add(item);
            }

            pptDataModel.EndToFront3Page.Description = DataConverter.ConvertListToTwoDArray(d1);
            pptDataModel.EndToFront3Page.Description2 = DataConverter.ConvertThreeElementListToTwoDArray(d2);

            #endregion

            #region ������2ҳ
            pptDataModel.EndToFront2Page = new EndToFront2();
            pptDataModel.EndToFront2Page.StructImage = new ImageModel
            {
                ImagePath = imgs.ElementAt(3).filePath,
                ImageName = imgs.ElementAt(3).name,
                Height = 300_0000,
                Width = 300_0000
            };

            string turnOnText = "";
            string turnOffText = "";
            foreach (var item in table.GetTurnOnData())
            {
                turnOnText += item + "\n";

            }
            foreach (var item in table.GetTurnOffData())
            {
                turnOffText += item + "\n";

            }
            pptDataModel.EndToFront2Page.TurnOn = turnOnText.TrimEnd('\n');
            pptDataModel.EndToFront2Page.TurnOff = turnOffText.TrimEnd('\n');
            #endregion


            #region ������1ҳ
            pptDataModel.LastPage = new LastPage();

            pptDataModel.LastPage.Image = new ImageModel
            {
                ImageName = imgs.ElementAt(4).name,
                ImagePath = imgs.ElementAt(4).filePath,
                Height = 500_0000,
                Width = 150_0000
            };
            pptDataModel.LastPage.Text1 = "Direct Mounting\n1.Typically, the die is mounted directly on the ground plane.\n2.If the thickness difference between the substrate (thickness c) and the die (thickness d) exceeds 0.05 mm (i.e., c �C d > 0.05 mm), it is recommended to first mount the die on a heat spreader, then attach the heat spreader to the ground plane.\n3.Heat Spreader Material: Molybdenum-copper (MoCu) alloy is commonly used.\n4.Heat Sink Thickness (b): Should be within the range of (c �C d �C 0.05 mm) to (c �Cd + 0.05 mm).\n5.Spacing (a): The gap between the bare die and the 50�� transmission line should typically be 0.05 mm to 0.1 mm. If the application frequency is higher than 40GHz, then this gap is recommended to be 0.05mm\n Wire Bonding Interconnection\nThe connection between the die and the 50�� transmission line is usually made using 25 \u03BCm diameter gold (Au) wires, bonded via wedge bonding or ball bonding processes.\nDie Attachment Methods\n1.Conductive Epoxy:\nAfter adhesive application, cure according to the manufacturer��s recommended temperature profile.\n2.Au-Sn80/20 Eutectic Bonding:\nUse preformed Au-Sn80/20 solder preforms.\nPerform bonding in an inert atmosphere (N\u2082 or forming gas: 90% N\u2082 + 10% H\u2082).\nKeep the time above 320\u2103 to less than 20 seconds to prevent excessive intermetallic formation.\n";
            pptDataModel.LastPage.Text2 = "Miller MMIC Inc. All rights reserved\nMiller MMIC, Inc. holds exclusive rights to the information presented in its Data Sheet and any accompanying materials. As a premier supplier of cutting-edge RF solutions, Miller MMIC has made this information easily accessible to its clients.\n\nAlthough Miller MMIC believes the information provided in its Data Sheet to be trustworthy, the company does not offer any guarantees as to its accuracy. Therefore, Miller MMIC bears no responsibility for the use of this information. It is worth mentioning that the information within the Data Sheet may be altered without prior notification.\n\nCustomers are encouraged to obtain and verify the most recent and pertinent information before placing any orders for Miller MMIC products. The information in the Data Sheet does not confer, either explicitly or implicitly, any rights or licenses with regards to patents or other forms of intellectual property to any third party.\n\nThe information provided in the Data Sheet, or its utilization, does not bestow any patent rights, licenses, or other forms of intellectual property rights to any individual or entity, whether in regards to the information itself or anything described by such information. Furthermore, Miller MMIC products are not intended for use as critical components in applications where failure could result in severe injury or death, such as medical or life-saving equipment, or life-sustaining applications, or in any situation where failure could cause serious personal injury or death.";
            #endregion


        }

        private void QueryDatabase(string pn, string sn, string startdatetime, string stopdatetime)
        {
            string connStr = "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";

            var repo = new TestRecordRepository(connStr);
            //var records = repo.GetRecordsByPN(
            //    pnList: new[] { "L004X"},
            //    startTime: new DateTime(2025, 10, 1),
            //    endTime: DateTime.Now
            //);
            string keyword = pn;
            var records = repo.GetRecordsByPN(
                keywords: new[] { keyword }
            );

            Console.WriteLine($"Found {records.Count} records.");
            foreach (var r in records)
            {
                Console.WriteLine($"{r.ID} | {r.PN} | {r.TestTime}");
            }

            string filePath = records.ElementAt(0).PN;

            var copier = new NetworkFolderCopier();

            // ��ѡ���Զ�����־���������д���ļ���
            // copier.Log = msg => File.AppendAllText("copy.log", $"{DateTime.Now:HH:mm:ss} {msg}\n");

            try
            {
                //����ON�����·��
                copier.CopyMatchingSubFolders(
                    networkRoot: @"\\DATAPC03\RFAutoTestReport$\Chip Verification",
                    PN: keyword,
                    ON: sn,
                    localTargetBase: System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiedReports"),
                    username: "",   // ���ձ�ʾʹ�õ�ǰ�û�ƾ��
                    password: ""
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"?? �����쳣: {ex.Message}");
            }

        }

        public void SetCurrentUserTaskManage(User user)
        {
            taskMangeControls.SetCurrentUser(user);
        }

        public void SetUsersNameLogWin(List<string> users)
        {
            logWin.SetComboxSourse(users);
        }
        private async void test3()
        {
            try
            {
                string pptFile = @"F:\PROJECT\ChipManualGeneration\exe\T_MML806_V3_Demo.pptx";
                string pdfFile = @"F:\PROJECT\ChipManualGeneration\exe\T_MML806_V3_Demo.pdf";

                // �ں�̨�߳�ִ�У����� UI ���ᣩ
                await Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));

                MessageBox.Show("PDF �����ɣ�", "�ɹ�", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ת��ʧ�ܣ�{ex.Message}", "����", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }


        private void GeneratePlotModels()
        {

            var finder = new TextFileFinder(
              rootDirectory: @"CopiedReports",
               extensions: new[] { ".txt", "s2p" }
              );
            //var finder = new TextFileFinder(); // ������ļ�������
            var allFiles = finder.FindAllTextFiles(); // �������·���б�

            // ��������ļ���������
            AmpfilierFilesbyGroup filesByGroup = AmplifierFileProcessor.ProcessFiles(allFiles);



        }
        private void test2()
        {
            curves.Clear();
            _plotNames.Clear();

            var condition = filter.GetFileterCondition();
            var filesByGroup = filter.FilesByGroup;
            if (condition.VD_VG_Conditon.Count > 1)
            {
                GeneratePlots(filesByGroup, condition);
                var temp = "Measurement Plots: S-parameters\n" + condition.VD_VG_Conditon.ElementAt(0).Replace('&', ',');
                _plotNames.Add(temp);
            }







            //if (filesByGroup.DataSparabyTemp.TryGetValue("25.0deg", out var s25))
            //{
            //    foreach (var item in s25)
            //    {
            //        foreach (var subCon in condition.VD_VG_Conditon)
            //        {
            //            if (subCon.Contains("&"))
            //            {
            //                string[] tmpArry = subCon.Split('&');
            //                if (item.Contains(tmpArry[0]))
            //                    filesModel25.SList.Add(item);
            //            }
            //            else
            //            {
            //                string[] tmpArry = subCon.Split(',');
            //                if (item.Contains(tmpArry[0]))
            //                    filesModel25.SList.Add(item);

            //            }
            //        }
            //    }



            //}

            //if (filesByGroup.PxdBbyTemp.TryGetValue("25.0deg", out var p1dbAt25))
            //{
            //    foreach (var item in p1dbAt25)
            //    {
            //        foreach (var subCon in condition.VD_VG_Conditon)
            //        {
            //            if (item.Contains(subCon))
            //                filesModel25.PxdbList.Add(item);
            //        }
            //    }

            //}


            //if (filesByGroup.OIP3byTemp.TryGetValue("25.0deg", out var oip3At25))
            //{
            //    foreach (var item in oip3At25)
            //    {
            //        foreach (var subCon in condition.VD_VG_Conditon)
            //        {
            //            if (item.Contains(subCon))
            //                filesModel25.OIP3List.Add(item);
            //        }
            //    }

            //}

            //if (filesByGroup.PsatbyTemp.TryGetValue("25.0deg", out var psatAt25))
            //{
            //    foreach (var item in psatAt25)
            //    {
            //        foreach (var subCon in condition.VD_VG_Conditon)
            //        {
            //            if (item.Contains(subCon))
            //                filesModel25.PsatList.Add(item);
            //        }
            //    }

            //}


            //if (filesByGroup.NFbyTemp.TryGetValue("25.0deg", out var nfAt25))
            //{
            //    foreach (var item in nfAt25)
            //    {
            //        foreach (var subCon in condition.VD_VG_Conditon)
            //        {
            //            if (item.Contains(subCon))
            //                filesModel25.NFList.Add(item);
            //        }
            //    }

            //}

            Temperature25FilePathModel filesModel25 = new Temperature25FilePathModel();
            const string tempKey = "25.0deg";
            var conditions = condition.VD_VG_Conditon;

            // 1. ��������Ҫ��������Լ����Ӧ��Ŀ���б�
            // ʹ��Ԫ��������ӳ����Դ�ֵ��Ŀ���б�
            var fileMappings = new List<(
                Dictionary<string, List<string>> sourceDict,
                List<string> targetList)>
    {
        (filesByGroup.DataSparabyTemp, filesModel25.SList),
        (filesByGroup.PxdBbyTemp, filesModel25.PxdbList),
        (filesByGroup.OIP3byTemp, filesModel25.OIP3List),
        (filesByGroup.PsatbyTemp, filesModel25.PsatList),
        (filesByGroup.NFbyTemp, filesModel25.NFList)
    };
            // 2. ��������ӳ�䲢ͳһ����
            foreach (var mapping in fileMappings)
            {
                // ���Դӵ�ǰ��Դ�ֵ��л�ȡ�ļ��б�
                if (mapping.sourceDict.TryGetValue(tempKey, out var files))
                {
                    // ʹ��ͨ�õ�ɸѡ������ȡƥ����ļ�
                    var matchingFiles = GetMatchingFiles(files, conditions);

                    // �����һ������ӵ�Ŀ���б���
                    mapping.targetList.AddRange(matchingFiles);
                }
            }


            GeneratePlotsByTemperaturen(filesModel25, condition.Min, condition.Max, LegendTextType.Elec);
            var tem1 = "Measurement Plots: S-parameters\n" + "TA = +25\u2103";
            var tem2 = "Measurement Plots: P1dB\n" + "TA = +25\u2103";
            var tem3 = "Measurement Plots: OIP3\n" + "TA = +25\u2103";
            var tem4 = "Measurement Plots: Psat\n" + "TA = +25\u2103";
            var tem5 = "Measurement Plots: Noise Figure\n" + "TA = +25\u2103";
            _plotNames.Add(tem1);
            _plotNames.Add(tem2);
            _plotNames.Add(tem3);
            _plotNames.Add(tem4);
            _plotNames.Add(tem5);

            List<Temperature25FilePathModel> filesModelVgs = new List<Temperature25FilePathModel>();

            //foreach (var item in filesByGroup.)
            //{ 

            //}

            foreach (var subCon in condition.VD_VG_Conditon)
            {
                string vd = "";
                if (subCon.Contains("&"))
                {
                    string[] tmpArry = subCon.Split('&');
                    vd = tmpArry[0];
                }
                else
                {
                    string[] tmpArry = subCon.Split(',');
                    vd = tmpArry[0];

                }
                var filesModelvg = new Temperature25FilePathModel();
                if (filesByGroup.DataSparabyVD.TryGetValue(vd, out var sVg))
                {
                    foreach (var item in sVg)
                    {
                        filesModelvg.SList.Add(item);
                    }
                }
                if (filesByGroup.PxdBbyVD.TryGetValue(vd, out var pxdbVg))
                {
                    foreach (var item in pxdbVg)
                    {
                        filesModelvg.PxdbList.Add(item);
                    }
                }
                if (filesByGroup.OIP3byVD.TryGetValue(vd, out var oip3Vg))
                {
                    foreach (var item in oip3Vg)
                    {
                        filesModelvg.OIP3List.Add(item);
                    }
                }
                if (filesByGroup.PsatbyVD.TryGetValue(vd, out var psatVg))
                {
                    foreach (var item in psatVg)
                    {
                        filesModelvg.PsatList.Add(item);
                    }
                }
                if (filesByGroup.NFbyVD.TryGetValue(vd, out var nfVg))
                {
                    foreach (var item in nfVg)
                    {
                        filesModelvg.NFList.Add(item);
                    }
                }


                var actually = new Temperature25FilePathModel();
                actually.SList = SelectUnique5VFilePerTemperature(filesModelvg.SList,vd);
                actually.PxdbList = SelectUnique5VFilePerTemperature(filesModelvg.PxdbList,vd);
                actually.OIP3List = SelectUnique5VFilePerTemperature(filesModelvg.OIP3List,vd);
                actually.PsatList = SelectUnique5VFilePerTemperature(filesModelvg.PsatList,vd);
                actually.NFList = SelectUnique5VFilePerTemperature(filesModelvg.NFList,vd);
                filesModelVgs.Add(actually);
            }
            foreach (var item in filesModelVgs)// ��ͬ��VD�£� ��ͬ�����¶�ͼ
            {
                if (item.SList.Count > 0)
                { 
                    List<string> culFiles = new List<string>();
                    culFiles.Add(item.SList.ElementAt(0));
                    culFiles.Add(item.PxdbList.ElementAt(0));
                    culFiles.Add(item.OIP3List.ElementAt(0));
                    culFiles.Add(item.PsatList.ElementAt(0));
                    culFiles.Add(item.NFList.ElementAt(0));
                    CalcuteParameter(culFiles, condition.FreqBands);

                    break;
                }
            }


               for(int i=0; i< filesModelVgs.Count; i++)// ��ͬ��VD�£� ��ͬ�����¶�ͼ
            {
                if (filesModelVgs.ElementAt(i).SList.Count > 0)
                {
                    GeneratePlotsByTemperaturen(filesModelVgs.ElementAt(i), condition.Min, condition.Max, LegendTextType.Temp);
                    // ����������ڴ�������б� �����������

                    var tem11 = "Measurement Plots: S-parameters\n" + condition.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                    var tem22 = "Measurement Plots: P1dB\n" + condition.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                    var tem33 = "Measurement Plots: OIP3\n" + condition.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                    var tem44 = "Measurement Plots: Psat\n" + condition.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                    var tem55 = "Measurement Plots: Noise Figure\n" + condition.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
      
                    _plotNames.Add(tem11);
                    _plotNames.Add(tem22);
                    _plotNames.Add(tem33);
                    _plotNames.Add(tem44);
                    _plotNames.Add(tem55);
                    //CalcuteParameter(List<string> files, List<string> freqBands)
                }


            }


            //if (ampParams.PsatbyVD.TryGetValue("-VD=5V&ID=90mA", out var psatAt90mA))
            //{
            //    Console.WriteLine("\nPsat files at -VD=5V&ID=90mA:");
            //    psatAt90mA.ForEach(Console.WriteLine);
            //}

        }

        // ----------------------------------------------------------------------
        // ͨ�ø�������������������Ҫ�ָ�ƥ���������߼�
        // ----------------------------------------------------------------------
        private IEnumerable<string> GetMatchingFiles(
            IEnumerable<string> files,
            IEnumerable<string> conditions)
        {
            // ʹ�� LINQ ��������ƥ����ļ�
            return files.Where(item =>
                conditions.Any(subCon =>
                {
                    // ȷ���ָ������������ '&' ���� '&'�������� ','
                    char splitChar = subCon.Contains('&') ? '&' : ',';

                    // �ָ��ַ���������ȡ��һ��������Ϊƥ������
                    // ʹ�� FirstOrDefault() ��� Split().FirstOrDefault()�����ܸ���
                    string matchPart = subCon.Split(splitChar).FirstOrDefault();

                    // ��� matchPart ��Ϊ���� item���ļ�·��������ƥ�䲿��
                    return !string.IsNullOrEmpty(matchPart) && item.Contains(matchPart);
                }));
        }
        public List<string> SelectUnique5VFilePerTemperature(List<string> allFiles, string vd)
        {
            var selectedFiles = allFiles
                // 1. Where (ɸѡ)��ֻ���� VD=5V ���ļ�
                .Where(filePath => ExtractVoltageKey(filePath) == vd)

                // 2. GroupBy (����)�����¶ȣ����� "25.0deg"�����з���
                .GroupBy(filePath => ExtractTemperatureKey(filePath))

                // 3. Select (ѡ��)����ÿ���¶�����ѡ���һ���ļ�·��
                .Select(group => group.First())

                // 4. ToList��ת��Ϊ�����б�
                .ToList();

            return selectedFiles;
        }

        // 1. ��ȡ�¶ȼ� (���� "25.0deg", "-40.0deg")
        private string ExtractTemperatureKey(string filePath)
        {
            // �����¶��������ļ���ĩβ�� "_XX.Xdeg_" ��ʽ֮ǰ
            // ʹ��������ʽ�������֡�С���㡢��ѡ�ĸ��ź� "deg"
            var match = Regex.Match(filePath, @"_(-?\d+\.\d+deg)_");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return "UNKNOWN_TEMP";
        }

        // 2. ��ȡ��ѹ�� (���� "VD=5V")
        private string ExtractVoltageKey(string filePath)
        {


            // �ҵ�·���е� VD/ID �Σ�Ȼ����ȡ VD ��ֵ��
            string normalizedPath = filePath.Replace('\\', '/');
            string[] pathSegments = normalizedPath.Split('/');

            string vdIdSegment = pathSegments.FirstOrDefault(s => s.StartsWith("-VD="));

            if (vdIdSegment != null)
            {
                // ʹ��������ȡ VD ��ֵ������� "-VD=5V&ID=90mA" ����ȡ "5V"
                var match = Regex.Match(vdIdSegment, @"-VD=(\d+V)");
                //if (match.Success)
                //{
                //    return match.Groups[1].Value; // ��ȡ "5V", "4V" ��
                //}
                //if (match.Success)
                //{
                //    return match.Groups[1].Value; // ��ȡ "5V", "4V" ��
                //}
                string str = match.ToString();
                return str;
            }
            return "UNKNOWN_V";
        }
        private void CureveGenerateLengdText(string filePath, out string temperature, out string elecParam)
        {
            string baseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string[] parts = baseName.Split('_');

            // ��ȡ�¶ȣ��� "deg" ��β�ĶΣ�
            temperature = parts.FirstOrDefault(p => p.EndsWith("deg", StringComparison.OrdinalIgnoreCase))
                              ?? "UnknownTemp";

            // ��ȡ�������������� "VD=" �ĶΣ�
            elecParam = parts.FirstOrDefault(p => p.Contains("VD="))
                            ?? "UnknownParam";
        }
        private void GeneratePlots(AmpfilierFilesbyGroup ampParams)
        {
            string filePath = "";

            #region ��ȡ��׼S����

            // ��ñ�׼S����
            if (ampParams.DataSparaFilePaths.Count > 4)
            {
                filePath = ampParams.DataSparaFilePaths[3];
            }
            else
            {

                filePath = ampParams.DataSparaFilePaths[0];
            }

            var analyzer = new S2PParser();
            analyzer.Parse(filePath);
            var point = curves.SPGenerateXYPointData(analyzer.S11, 0);
            string yLable = "INPUT RETURN LOSS(dB)";
            string xLable = "FREQUENCY(GHz)";
            string title = "";
            string temperature = "";
            string elecParam = "";
            CureveGenerateLengdText(filePath, out temperature, out elecParam);
            string legend = elecParam;
            curves.AddPlot(curves.GeneratePlotParameters(point, xLable, yLable, title, legend));

            point = curves.SPGenerateXYPointData(analyzer.S12, 0);
            yLable = "ISOLATION(dB)";
            curves.AddPlot(curves.GeneratePlotParameters(point, xLable, yLable, title, legend));

            point = curves.SPGenerateXYPointData(analyzer.S21, 0);
            yLable = "Gain(dB)";
            curves.AddPlot(curves.GeneratePlotParameters(point, xLable, yLable, title, legend));

            point = curves.SPGenerateXYPointData(analyzer.S22, 0);
            yLable = "OUTPUT RETURN LOSS(dB)";
            curves.AddPlot(curves.GeneratePlotParameters(point, xLable, yLable, title, legend));
            #endregion


            #region 25������s����
            var points = new Collection<XYPoint>();
            var legends = new Collection<string>();
            if (ampParams.DataSparabyTemp.TryGetValue("25.0deg", out var s2pAt25))
            {
                foreach (var item in s2pAt25)
                {
                    analyzer.Parse(item);
                    var pointTmp = curves.SPGenerateXYPointData(analyzer.S11, 0);
                    CureveGenerateLengdText(item, out temperature, out elecParam);
                    points.Add(pointTmp);
                    legends.Add(elecParam);
                }
                //s2pAt25.ForEach(Console.WriteLine);

            }
            filter.AddLevels(legends.ToList());
            yLable = "INPUT RETURN LOSS(dB)";
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));


            points.Clear();
            legends.Clear();
            yLable = "ISOLATION(dB)";
            foreach (var item in s2pAt25)
            {
                analyzer.Parse(item);
                var pointTmp = curves.SPGenerateXYPointData(analyzer.S12, 0);
                CureveGenerateLengdText(item, out temperature, out elecParam);
                points.Add(pointTmp);
                legends.Add(elecParam);
            }
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));



            points.Clear();
            legends.Clear();
            yLable = "Gain(dB)";
            foreach (var item in s2pAt25)
            {
                analyzer.Parse(item);
                var pointTmp = curves.SPGenerateXYPointData(analyzer.S21, 0);
                CureveGenerateLengdText(item, out temperature, out elecParam);
                points.Add(pointTmp);
                legends.Add(elecParam);
            }
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));



            points.Clear();
            legends.Clear();
            yLable = "Gain(dB)";
            foreach (var item in s2pAt25)
            {
                analyzer.Parse(item);
                var pointTmp = curves.SPGenerateXYPointData(analyzer.S22, 0);
                CureveGenerateLengdText(item, out temperature, out elecParam);
                points.Add(pointTmp);
                legends.Add(elecParam);
            }
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));
            #endregion

            #region 25��nf psat pxdb�� oip3

            points.Clear();
            legends.Clear();
            yLable = "P1dB(dBm)";
            var txtParser = new TextFileParser();
            //textParser.ParsePsatFile(ampParams.NFFiles.ElementAt(0));
            if (ampParams.PxdBbyTemp.TryGetValue("25.0deg", out var p1dbAt25))
            {
                foreach (var item in p1dbAt25)
                {
                    txtParser.Parse(item);
                    var pointTmp = curves.SPGenerateXYPointData(txtParser.Points);
                    CureveGenerateLengdText(item, out temperature, out elecParam);
                    points.Add(pointTmp);
                    legends.Add(elecParam);
                }
                //s2pAt25.ForEach(Console.WriteLine);
            }
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));




            points.Clear();
            legends.Clear();
            yLable = "OUTPUT IP3(dBm)";

            //textParser.ParsePsatFile(ampParams.NFFiles.ElementAt(0));
            if (ampParams.OIP3byTemp.TryGetValue("25.0deg", out var ip3At25))
            {
                foreach (var item in ip3At25)
                {
                    txtParser.Parse(item);
                    var pointTmp = curves.SPGenerateXYPointData(txtParser.Points);
                    CureveGenerateLengdText(item, out temperature, out elecParam);
                    points.Add(pointTmp);
                    legends.Add(elecParam);
                }
                //s2pAt25.ForEach(Console.WriteLine);
            }
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));



            points.Clear();
            legends.Clear();
            yLable = "Psat(dBm)";

            //textParser.ParsePsatFile(ampParams.NFFiles.ElementAt(0));
            if (ampParams.PsatbyTemp.TryGetValue("25.0deg", out var psatAt25))
            {
                foreach (var item in psatAt25)
                {
                    txtParser.Parse(item);
                    var pointTmp = curves.SPGenerateXYPointData(txtParser.Points);
                    CureveGenerateLengdText(item, out temperature, out elecParam);
                    points.Add(pointTmp);
                    legends.Add(elecParam);
                }
                //s2pAt25.ForEach(Console.WriteLine);
            }
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));


            points.Clear();
            legends.Clear();
            yLable = "NOISE FIGURE(dBm)";

            //textParser.ParsePsatFile(ampParams.NFFiles.ElementAt(0));
            if (ampParams.NFbyTemp.TryGetValue("25.0deg", out var nfAt25))
            {
                foreach (var item in nfAt25)
                {
                    txtParser.Parse(item);
                    var pointTmp = curves.SPGenerateXYPointData(txtParser.Points);
                    CureveGenerateLengdText(item, out temperature, out elecParam);
                    points.Add(pointTmp);
                    legends.Add(elecParam);
                }
                //s2pAt25.ForEach(Console.WriteLine);
            }
            curves.AddPlot(curves.GeneratePlotParameters(points, xLable, yLable, title, legends));

            #endregion



            //var analyzer2 = new S2PParser();
            //analyzer2.Parse(@"F:\PROJECT\ChipManualGeneration\ԭʼ����\MML004X_V2-����\MML004X_V2-25\-VD=3V&ID=43mA\L004X��L024X��L026X_MML004X_V2-25_-VD=3V&ID=43mA_2025-09-01 14.49.13_25.0deg_SPara.s2p");

            //var analyzer3 = new S2PParser();
            //analyzer3.Parse(@"F:\PROJECT\ChipManualGeneration\ԭʼ����\MML004X_V2-����\MML004X_V2-25\-VD=5V&ID=90mA\L004X��L024X��L026X_MML004X_V2-25_-VD=5V&ID=90mA_2025-09-01 15.05.38_25.0deg_SPara.s2p");


            //var points = new Collection<XYPoint>();
            //var point1 = SPGenerateXYPointData(analyzer.S11, 0);
            //var point2 = SPGenerateXYPointData(analyzer2.S11, 0);
            //var point3 = SPGenerateXYPointData(analyzer3.S11, 0);

            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "INPUT RETURN LOSS(dB)";
            //var legends = new Collection<string>();
            //legends.Add("-VD=4V&ID=67mA");
            //legends.Add("-VD=3V&ID=43mA");
            //legends.Add("-VD=5V&ID=90mA");
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //points.Clear();
            //point1 = SPGenerateXYPointData(analyzer.S12, 0);
            //point2 = SPGenerateXYPointData(analyzer2.S12, 0);
            //point3 = SPGenerateXYPointData(analyzer3.S12, 0);
            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "ISOLATION(dB)";
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //points.Clear();
            //point1 = SPGenerateXYPointData(analyzer.S21, 0);
            //point2 = SPGenerateXYPointData(analyzer2.S21, 0);
            //point3 = SPGenerateXYPointData(analyzer3.S21, 0);
            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "GAIN(dB)";
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //points.Clear();
            //point1 = SPGenerateXYPointData(analyzer.S22, 0);
            //point2 = SPGenerateXYPointData(analyzer2.S22, 0);
            //point3 = SPGenerateXYPointData(analyzer3.S22, 0);
            //points.Add(point1);
            //points.Add(point2);
            //points.Add(point3);
            //yLable = "OUTPUT RETURN LOSS(dB)";
            //AddPlot(GeneratePlotParameters(points, xLable, yLable, title, legends));

            //int index = 0;
            // �洢ͼƬ��
            //foreach (var item in vm.Plots)
            //{
            //    //item.Refresh();
            //    //Console.WriteLine(vm.Plots.Count);
            //    //string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            //    //string fileName = $"myplot_{index}.png";
            //    //item.Plot.SavePng(fileName, 600, 500);
            //    //item.Refresh();

            //    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            //    string folder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pic");
            //    Directory.CreateDirectory(folder);
            //    string fileName = System.IO.Path.Combine(folder, $"{index}.png");

            //    Application.Current.Dispatcher.Invoke(() =>
            //    {
            //        item.Plot.SavePng(fileName, 600, 500);
            //    });
            //    index++;
            //}


        }
        //private void btn_Generate_Click(object sender, RoutedEventArgs e)
        private void GeneratePlots(AmpfilierFilesbyGroup ampParams, FileterConditionModel condition)
        {

            string filePath = "";

            //#region ��ȡ��׼S����

            // ��ñ�׼S����
            if (ampParams.DataSparaFilePaths.Count > 4)
            {
                filePath = ampParams.DataSparaFilePaths[3];
            }
            else
            {

                filePath = ampParams.DataSparaFilePaths[0];
            }
            //curves.Clear();
            //��ǰ�׶�ֱ�Ӵ����һ��
            //List<string> filesArry = new List<string>();
            //if( ampParams.DataSparabyTemp.TryGetValue("25.0deg",out var  files))
            //{

            //    foreach (var item in files)
            //    {
            //        foreach (var subCon in condition.VD_VG_Conditon)
            //        {
            //            if (subCon.Contains("&"))
            //            {
            //                string[] tmpArry = subCon.Split('&');
            //                if (item.Contains(tmpArry[0]))
            //                    filesArry.Add(item);
            //            }
            //            else
            //            {
            //                string[] tmpArry = subCon.Split(',');
            //                if (item.Contains(tmpArry[0]))
            //                    filesArry.Add(item);

            //            }


            //            //string[] tmpArry = subCon.Split("");

            //        }
            //    }

            //}
            //List<string> files2 = new List<string>();
            //files2.Add(filesArry.ElementAt(0));
            //GenerateSParaPlots(files2, condition.Min, condition.Max, LegendTextType.Elec);

            if (!ampParams.DataSparabyTemp.TryGetValue("25.0deg", out var files))
            {
                // ��� files Ϊ�ջ�δ�ҵ������򷵻ػ��׳��쳣
                return; // ���� throw new KeyNotFoundException("δ�ҵ� '25.0deg' �����ݡ�");
            }

            // 1. ʹ�� LINQ ���ҵ�һ��ƥ����ļ�·��
            string firstMatchingFile = files.FirstOrDefault(item =>
            {
                // ���������������ҵ���һ�����ϵ��ļ�������
                return condition.VD_VG_Conditon.Any(subCon =>
                {
                    // �Ż��㣺ʹ�� char.Split ���� string.Split
                    char splitChar = subCon.Contains('&') ? '&' : ',';

                    // �ָ��ַ���������ȡ��һ��������Ϊƥ������
                    // �Ż���ʹ�� subCon.Split(splitChar).FirstOrDefault() ���洴�����������ٷ��� [0]
                    string matchPart = subCon.Split(splitChar).FirstOrDefault();

                    // ��� item���ļ�·�����Ƿ����ƥ�䲿��
                    return matchPart != null && item.Contains(matchPart);
                });
            });


            // 2. ����Ƿ��ҵ��ļ��������л�ͼ
            if (firstMatchingFile != null)
            {
                // �������ļ����� List<string> (files2)
                List<string> files2 = new List<string> { firstMatchingFile };

                // ִ�л�ͼ����
                GenerateSParaPlots(files2, condition.Min, condition.Max, LegendTextType.Elec);
            }

        }

        private void GenerateSParaPlots(List<string> files, double xMin, double xMax, LegendTextType type)
        {

            var plotS21 = new PlotModel();
            plotS21.YLabel = "GAIN(DB)";
            plotS21.XLabel = "FREQUENCY(GHz)";


            var plotS11 = new PlotModel();
            plotS11.YLabel = "INPUT RETURN LOSS(dB)";
            plotS11.XLabel = "FREQUENCY(GHz)";

            var plotS12 = new PlotModel();
            plotS12.YLabel = "ISOLATION(dB)";
            plotS12.XLabel = "FREQUENCY(GHz)";

            var plotS22 = new PlotModel();
            plotS22.YLabel = "OUTPUT RETURN LOSS(dB)";
            plotS22.XLabel = "FREQUENCY(GHz)";

            var analyzer = new S2PParser();

            foreach (var file in files)
            {
                //var analyzer = new S2PParser();
                analyzer.Parse(file);

                var pointS11 = curves.SPGenerateXYPointData(analyzer.S11, 0);
                var pointS12 = curves.SPGenerateXYPointData(analyzer.S12, 0);
                var pointS21 = curves.SPGenerateXYPointData(analyzer.S21, 0);
                var pointS22 = curves.SPGenerateXYPointData(analyzer.S22, 0);
                string temperature = "";
                string elecParam = "";
                CureveGenerateLengdText(file, out temperature, out elecParam);
                var curveS11 = new CurevelModel();
                var curveS12 = new CurevelModel();
                var curveS21 = new CurevelModel();
                var curveS22 = new CurevelModel();
                curveS11.XData = pointS11.XArrys;
                curveS11.YData = pointS11.YArrys;

                curveS12.XData = pointS12.XArrys;
                curveS12.YData = pointS12.YArrys;

                curveS21.XData = pointS21.XArrys;
                curveS21.YData = pointS21.YArrys;

                curveS22.XData = pointS22.XArrys;
                curveS22.YData = pointS22.YArrys;
                switch (type)
                {
                    case LegendTextType.Temp:// legend ���¶ȵ���ͬ
                        curveS11.Legend = temperature;
                        curveS12.Legend = temperature;
                        curveS21.Legend = temperature;
                        curveS22.Legend = temperature;
                        break;

                    case LegendTextType.Elec:// �Ե�����������ͬ
                        curveS11.Legend = elecParam.Replace("&", ",");
                        curveS12.Legend = elecParam.Replace("&", ",");
                        curveS21.Legend = elecParam.Replace("&", ",");
                        curveS22.Legend = elecParam.Replace("&", ",");
                        break;
                }

                plotS11.Cureves.Add(curveS11);
                plotS12.Cureves.Add(curveS12);
                plotS21.Cureves.Add(curveS21);
                plotS22.Cureves.Add(curveS22);



            }

            plotS11.SetYAxisLimits();
            plotS12.SetYAxisLimits();
            plotS21.SetYAxisLimits();
            plotS22.SetYAxisLimits();


            plotS11.xMin = xMin;
            plotS11.xMax = xMax;
            plotS12.xMin = xMin;
            plotS12.xMax = xMax;
            plotS21.xMin = xMin;
            plotS21.xMax = xMax;
            plotS22.xMin = xMin;
            plotS22.xMax = xMax;


            //���µ����� ��x����ߣ���Сֵ������߿̶��ߴ��� ���ֵ�����ұ߿̶��߳��� ���õķ����ǵ���ÿ�������� ѡ��10������������
            if (PlotModel.CalculateFixedInterval((int)xMin, (int)xMax, 10, out int xInterval))
            {
                plotS11.xAxisInterval = xInterval;
                plotS12.xAxisInterval = xInterval;
                plotS21.xAxisInterval = xInterval;
                plotS22.xAxisInterval = xInterval;
            }
            double newYMin, newYMax;
            int yInterval;
            const int TargetDivisions = 10; // Ŀ�� 10 ���̶�

            PlotModel.CalculateNiceRange((double)plotS11.yMin, (double)plotS11.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS11.yMin = newYMin;
            plotS11.yMax = newYMax;
            plotS11.yAxisInterval = yInterval;
            plotS11.Alignment = ScottPlot.Alignment.UpperRight;
            //plotS11.Alignment = ScottPlot.Alignment.UpperRight;


            PlotModel.CalculateNiceRange((double)plotS12.yMin, (double)plotS12.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS12.yMin = newYMin;
            plotS12.yMax = newYMax;
            plotS12.yAxisInterval = yInterval;
            //plotS12.Alignment = ScottPlot.Alignment.UpperRight;


            PlotModel.CalculateNiceRange((double)plotS21.yMin, (double)plotS21.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS21.yMin = newYMin;
            plotS21.yMax = newYMax;
            plotS21.yAxisInterval = yInterval;
            plotS21.Alignment = ScottPlot.Alignment.LowerRight;
            //plotS11.Alignment = ScottPlot.Alignment.LowerRight;

            PlotModel.CalculateNiceRange((double)plotS22.yMin, (double)plotS22.yMax, TargetDivisions,
                                                 out newYMin, out newYMax, out yInterval);
            plotS22.yMin = newYMin;
            plotS22.yMax = newYMax;
            plotS22.yAxisInterval = yInterval;
            plotS22.Alignment = ScottPlot.Alignment.UpperRight;


            curves.AddPlot(plotS21);
            curves.AddPlot(plotS11);
            curves.AddPlot(plotS12);
            curves.AddPlot(plotS22);
        }

        /// <summary>
        /// ��������NF, Pxdb, OIP3, Psat�⼸��ͼ��
        /// </summary>
        /// <param name="files"></param>
        /// <param name="xMin"></param>
        /// <param name="xMax"></param>
        /// <param name="type"></param>
        private void GeneratePlotsByTemperaturen(Temperature25FilePathModel filesModel25, double xMin, double xMax, LegendTextType type)

        {

            
            #region S��������
            ////////////s11
            var plotS11 = new PlotModel();
            plotS11.YLabel = "INPUT RETURN LOSS(dB)";
            plotS11.XLabel = "FREQUENCY(GHz)";

            ////////////s12
            var plotS12 = new PlotModel();
            plotS12.YLabel = "ISOLATION(dB)";
            plotS12.XLabel = "FREQUENCY(GHz)";

            ////////////s21
            var plotS21 = new PlotModel();
            plotS21.YLabel = "GAIN(DB)";
            plotS21.XLabel = "FREQUENCY(GHz)";

            ////////////s22
            var plotS22 = new PlotModel();
            plotS22.YLabel = "OUTPUT RETURN LOSS(dB)";
            plotS22.XLabel = "FREQUENCY(GHz)";
            var analyzer = new S2PParser();

            foreach (var file in filesModel25.SList)
            {
                //var analyzer = new S2PParser();
                analyzer.Parse(file);

                var pointS11 = curves.SPGenerateXYPointData(analyzer.S11, 0);
                var pointS12 = curves.SPGenerateXYPointData(analyzer.S12, 0);
                var pointS21 = curves.SPGenerateXYPointData(analyzer.S21, 0);
                var pointS22 = curves.SPGenerateXYPointData(analyzer.S22, 0);
                string temperature = "";
                string elecParam = "";
                CureveGenerateLengdText(file, out temperature, out elecParam);
                var curveS11 = new CurevelModel();
                var curveS12 = new CurevelModel();
                var curveS21 = new CurevelModel();
                var curveS22 = new CurevelModel();
                curveS11.XData = pointS11.XArrys;
                curveS11.YData = pointS11.YArrys;

                curveS12.XData = pointS12.XArrys;
                curveS12.YData = pointS12.YArrys;

                curveS21.XData = pointS21.XArrys;
                curveS21.YData = pointS21.YArrys;

                curveS22.XData = pointS22.XArrys;
                curveS22.YData = pointS22.YArrys;
                switch (type)
                {
                    case LegendTextType.Temp:// legend ���¶ȵ���ͬ
                        curveS11.Legend = temperature;
                        curveS12.Legend = temperature;
                        curveS21.Legend = temperature;
                        curveS22.Legend = temperature;
                        break;

                    case LegendTextType.Elec:// �Ե�����������ͬ
                        curveS11.Legend = elecParam.Replace("&", ",");
                        curveS12.Legend = elecParam.Replace("&", ",");
                        curveS21.Legend = elecParam.Replace("&", ",");
                        curveS22.Legend = elecParam.Replace("&", ",");
                        break;
                }

                plotS11.Cureves.Add(curveS11);
                plotS12.Cureves.Add(curveS12);
                plotS21.Cureves.Add(curveS21);
                plotS22.Cureves.Add(curveS22);



            }

            plotS11.SetYAxisLimits();
            plotS12.SetYAxisLimits();
            plotS21.SetYAxisLimits();
            plotS22.SetYAxisLimits();


            plotS11.xMin = xMin;
            plotS11.xMax = xMax;
            plotS12.xMin = xMin;
            plotS12.xMax = xMax;
            plotS21.xMin = xMin;
            plotS21.xMax = xMax;
            plotS22.xMin = xMin;
            plotS22.xMax = xMax;


            //���µ����� ��x����ߣ���Сֵ������߿̶��ߴ��� ���ֵ�����ұ߿̶��߳��� ���õķ����ǵ���ÿ�������� ѡ��10������������
            if (PlotModel.CalculateFixedInterval((int)xMin, (int)xMax, 10, out int xInterval))
            {
                plotS11.xAxisInterval = xInterval;
                plotS12.xAxisInterval = xInterval;
                plotS21.xAxisInterval = xInterval;
                plotS22.xAxisInterval = xInterval;
            }
            double newYMin, newYMax;
            int yInterval;
            const int TargetDivisions = 10; // Ŀ�� 10 ���̶�

            PlotModel.CalculateNiceRange((double)plotS11.yMin, (double)plotS11.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS11.yMin = newYMin;
            plotS11.yMax = newYMax;
            plotS11.yAxisInterval = yInterval;
            plotS11.Alignment = ScottPlot.Alignment.UpperRight;


            PlotModel.CalculateNiceRange((double)plotS12.yMin, (double)plotS12.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS12.yMin = newYMin;
            plotS12.yMax = newYMax;
            plotS12.yAxisInterval = yInterval;
            //plotS12.Alignment = ScottPlot.Alignment.UpperRight;


            PlotModel.CalculateNiceRange((double)plotS21.yMin, (double)plotS21.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS21.yMin = newYMin;
            plotS21.yMax = newYMax;
            plotS21.yAxisInterval = yInterval;
            plotS21.Alignment = ScottPlot.Alignment.LowerRight;


            PlotModel.CalculateNiceRange((double)plotS22.yMin, (double)plotS22.yMax, TargetDivisions,
                                                 out newYMin, out newYMax, out yInterval);
            plotS22.yMin = newYMin;
            plotS22.yMax = newYMax;
            plotS22.yAxisInterval = yInterval;
            plotS22.Alignment = ScottPlot.Alignment.UpperCenter;


            curves.AddPlot(plotS21);
            curves.AddPlot(plotS11);
            curves.AddPlot(plotS12);
            curves.AddPlot(plotS22);

            #endregion


            #region ������������
            ////////////pxdb
            var plotP1db = new PlotModel();
            plotP1db.YLabel = "P1dB(dBm)";
            plotP1db.XLabel = "FREQUENCY(GHz)";
            var txtParser = new TextFileParser();
            //textParser.ParsePsatFile(ampParams.NFFiles.ElementAt(0));

            foreach (var file in filesModel25.PxdbList)
            {
                txtParser.Parse(file);
                var point = curves.SPGenerateXYPointData(txtParser.Points);
                string temperature = "";
                string elecParam = "";
                CureveGenerateLengdText(file, out temperature, out elecParam);
                var curve = new CurevelModel();
                curve.XData = point.XArrys;
                curve.YData = point.YArrys;

                switch (type)
                {
                    case LegendTextType.Temp:
                        curve.Legend = temperature;
                        break;
                    case LegendTextType.Elec:
                        curve.Legend = elecParam.Replace("&", ",");
                        break;
                }
                plotP1db.Cureves.Add(curve);
            }
            plotP1db.SetYAxisLimits();
            plotP1db.xMin = xMin;
            plotP1db.xMax = xMax;
            plotP1db.xAxisInterval = xInterval;


            PlotModel.CalculateNiceRange((double)plotP1db.yMin, (double)plotP1db.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotP1db.yMin = newYMin;
            plotP1db.yMax = newYMax;
            plotP1db.yAxisInterval = yInterval;


            ////////////oip3
            var plotOIP3 = new PlotModel();
            plotOIP3.YLabel = "OUTPUT IP3(dBm)";
            plotOIP3.XLabel = "FREQUENCY(GHz)";


            foreach (var file in filesModel25.OIP3List)
            {
                txtParser.Parse(file);
                var point = curves.SPGenerateXYPointData(txtParser.Points);
                string temperature = "";
                string elecParam = "";
                CureveGenerateLengdText(file, out temperature, out elecParam);
                var curve = new CurevelModel();
                curve.XData = point.XArrys;
                curve.YData = point.YArrys;

                switch (type)
                {
                    case LegendTextType.Temp:
                        curve.Legend = temperature;
                        break;
                    case LegendTextType.Elec:
                        curve.Legend = elecParam.Replace("&", ",");
                        break;
                }
                plotOIP3.Cureves.Add(curve);
            }
            plotOIP3.SetYAxisLimits();
            plotOIP3.xMin = xMin;
            plotOIP3.xMax = xMax;
            plotOIP3.xAxisInterval = xInterval;


            PlotModel.CalculateNiceRange((double)plotOIP3.yMin, (double)plotOIP3.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotOIP3.yMin = newYMin;
            plotOIP3.yMax = newYMax;
            plotOIP3.yAxisInterval = yInterval;


            ////////////psat
            var plotPsat = new PlotModel();
            plotPsat.YLabel = "Psat(dBm)";
            plotPsat.XLabel = "FREQUENCY(GHz)";
            foreach (var file in filesModel25.PsatList)
            {
                txtParser.Parse(file);
                var point = curves.SPGenerateXYPointData(txtParser.Points);
                string temperature = "";
                string elecParam = "";
                CureveGenerateLengdText(file, out temperature, out elecParam);
                var curve = new CurevelModel();
                curve.XData = point.XArrys;
                curve.YData = point.YArrys;

                switch (type)
                {
                    case LegendTextType.Temp:
                        curve.Legend = temperature;
                        break;
                    case LegendTextType.Elec:
                        curve.Legend = elecParam.Replace("&", ",");
                        break;
                }
                plotPsat.Cureves.Add(curve);
            }
            plotPsat.SetYAxisLimits();
            plotPsat.xMin = xMin;
            plotPsat.xMax = xMax;
            plotPsat.xAxisInterval = xInterval;


            PlotModel.CalculateNiceRange((double)plotPsat.yMin, (double)plotPsat.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotPsat.yMin = newYMin;
            plotPsat.yMax = newYMax;
            plotPsat.yAxisInterval = yInterval;


            ////////////NF
            var plotNF = new PlotModel();
            plotNF.YLabel = "NOISE FIGURE(dBm)";
            plotNF.XLabel = "FREQUENCY(GHz)";
            foreach (var file in filesModel25.NFList)
            {
                txtParser.Parse(file);
                var point = curves.SPGenerateXYPointData(txtParser.Points);
                string temperature = "";
                string elecParam = "";
                CureveGenerateLengdText(file, out temperature, out elecParam);
                var curve = new CurevelModel();
                curve.XData = point.XArrys;
                curve.YData = point.YArrys;

                switch (type)
                {
                    case LegendTextType.Temp:
                        curve.Legend = temperature;
                        break;
                    case LegendTextType.Elec:
                        curve.Legend = elecParam.Replace("&", ",");
                        break;
                }
                plotNF.Cureves.Add(curve);
            }
            plotNF.SetYAxisLimits();
            plotNF.xMin = xMin;
            plotNF.xMax = xMax;
            plotNF.xAxisInterval = xInterval;


            PlotModel.CalculateNiceRange((double)plotNF.yMin, (double)plotNF.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotNF.yMin = newYMin;
            plotNF.yMax = newYMax;
            plotNF.yAxisInterval = yInterval;
            plotNF.Alignment = ScottPlot.Alignment.UpperCenter;

            curves.AddPlot(plotP1db);
            curves.AddPlot(plotOIP3);
            curves.AddPlot(plotPsat);
            curves.AddPlot(plotNF);
            #endregion
        }

        private void CalcuteParameter(List<string> files, List<string> freqBands)
        {
            // files 0-s�����ļ���1-pxdb��2-oip3��3-psat��4-nf
            var analyzer = new S2PParser();
            var txtParser = new TextFileParser();
            //textParser.ParsePsatFile(ampParams.NFFiles.ElementAt(0)); 
            
            var table1 = new List<List<double>>();
            var bands = new List<double>();
            var s11 = new List<double>();
            var gf = new List<double>();//Gain Flatness

            var s21 = new List<double>();
            var s22 = new List<double>();

            var pxdb = new List<double>();
            var oip3 = new List<double>();
            var psat = new List<double>();
            var nf = new List<double>();
            foreach (var band in freqBands)// ��������Ƶ�Σ� ����Ƶ�ε�min�� type�� amx
            {
                double min;
                double max;
                
                if (TryParseMinMaxBand(band, out min, out max))
                {
                    bands.Add(min);
                    bands.Add(0);
                    bands.Add(max);
                    analyzer.Parse(files.ElementAt(0));
                    var pointS11 = curves.SPGenerateXYPointData(analyzer.S11, 0);
                    var pointS12 = curves.SPGenerateXYPointData(analyzer.S12, 0);
                    var pointS21 = curves.SPGenerateXYPointData(analyzer.S21, 0);
                    var pointS22 = curves.SPGenerateXYPointData(analyzer.S22, 0);
                    var subPointS11 = FilterXYPointData(pointS11, min, max);
                    var subPointS12 = FilterXYPointData(pointS12, min, max);
                    var subPointS21 = FilterXYPointData(pointS21, min, max);
                    var subPointS22 = FilterXYPointData(pointS22, min, max);
                    double minS11 = subPointS11.YArrys.Min();
                    double maxS11 = subPointS11.YArrys.Max();
                    double typeS11 = (maxS11 + minS11) / 2.0;
                    double minS12 = subPointS12.YArrys.Min();
                    double maxS12 = subPointS12.YArrys.Max();
                    double typeS12 = (maxS12 + minS12) / 2.0;
                    double minS21 = subPointS21.YArrys.Min();
                    double maxS21 = subPointS21.YArrys.Max();
                    double type21 = (maxS21 + minS21) / 2.0;
                    double minS22 = subPointS22.YArrys.Min();
                    double maxS22 = subPointS22.YArrys.Max();
                    double type22 = (maxS22 + minS22) / 2.0;
                    s11.Add(minS11);
                    
                    s11.Add(typeS11);
                    s11.Add(maxS11);


                    int minIndex = Array.IndexOf(subPointS21.YArrys, minS21);
                    int maxIndex = Array.IndexOf(subPointS21.YArrys, maxS21);
                    gf.Add(0);
                    double value = (subPointS21.YArrys[minIndex] - subPointS21.YArrys[maxIndex])/(subPointS21.XArrys[minIndex]- subPointS21.XArrys[maxIndex]);
                    gf.Add(Math.Abs(value));
                    gf.Add(0);
                    

                    s21.Add(minS21);
                    s21.Add(type21);
                    s21.Add(maxS21);
                    

                    s22.Add(minS22);
                    s22.Add(type22);
                    s22.Add(maxS22);
                    


                    txtParser.Parse(files.ElementAt(1));
                    var pointP1db = curves.SPGenerateXYPointData(txtParser.Points);
                    var subPointP1db = FilterXYPointData(pointP1db, min, max);
                    double minP1db = subPointP1db.YArrys.Min();
                    double maxP1db = subPointP1db.YArrys.Max();
                    double typeP1db = (maxP1db + minP1db) / 2.0;
                    pxdb.Add(minP1db);
                    pxdb.Add(typeP1db);
                    pxdb.Add(maxP1db);
                    

                    txtParser.Parse(files.ElementAt(2));
                    var pointOIP3 = curves.SPGenerateXYPointData(txtParser.Points);
                    var subPointOIP3 = FilterXYPointData(pointOIP3, min, max);
                    double minOIP3 = subPointOIP3.YArrys.Min();
                    double maxOIP3 = subPointOIP3.YArrys.Max();
                    double typeOIP3 = (maxOIP3 + minOIP3) / 2.0;
                    oip3.Add(minOIP3);
                    oip3.Add(typeOIP3);
                    oip3.Add(maxOIP3);
                    

                    txtParser.Parse(files.ElementAt(3));
                    var pointPsat = curves.SPGenerateXYPointData(txtParser.Points);
                    var subPointPsat = FilterXYPointData(pointPsat, min, max);
                    double minPsat = subPointPsat.YArrys.Min();
                    double maxPsat = subPointPsat.YArrys.Max();
                    double typePsat = (maxPsat + minPsat) / 2.0;
                    psat.Add(minPsat);
                    psat.Add(typePsat);
                    psat.Add(maxPsat);
                    

                    txtParser.Parse(files.ElementAt(4));
                    var pointNF = curves.SPGenerateXYPointData(txtParser.Points);
                    var subPointNF = FilterXYPointData(pointNF, min, max);
                    double minNF = subPointNF.YArrys.Min();
                    double maxNF = subPointNF.YArrys.Max();
                    double typeNF = (maxNF + minNF) / 2.0;
                    nf.Add(minNF);
                    nf.Add(typeNF);
                    nf.Add(maxNF);
                    
                
                }

            }
            table1.Add(bands);
            table1.Add(s21);
            table1.Add(gf);
            table1.Add(nf);
            table1.Add(pxdb);
            table1.Add(psat);
            table1.Add(oip3);
            table1.Add(s11);
            table1.Add(s22); 
            table.SetParametersTableInfo(table1);
        }
        private double CalcuteAverageParameter(double[] value)
        {
            if (value == null || !value.Any())
            {
                return double.NaN;
            }

            return (value.Min() + value.Max()) / 2.0;

        }
        /// <param name="band">�����������ֵ��ַ������м����������ָ�����</param>
        /// <param name="minValue">���������������Сֵ��</param>
        /// <param name="maxValue">����������������ֵ��</param>
        /// <returns>����ɹ��������������֣��򷵻� true�����򷵻� false��</returns>
        public bool TryParseMinMaxBand(string band, out double minValue, out double maxValue)
        {
            minValue = 0.0;
            maxValue = 0.0;

            if (string.IsNullOrWhiteSpace(band))
            {
                return false;
            }

            // �Ż����������ʽ��ֻƥ��Ǹ��������ֿ�ѡ��С���㣩
            // �ؼ��ı䣺�Ƴ���ǰ��� "-?"
            // \d+ : ƥ��һ���������� (0-9)
            // \.? : ƥ���ѡ��С����
            // \d* : ƥ�� 0 ���������֣�����֧�� .5, 10.0 �ȣ�
            const string pattern = @"\d+\.?\d*";

            // 1. ��������ƥ��Ǹ����ֵĲ���
            // ע�⣺��������ۺš����ŵ�����ֻ����Ϊ�ָ���
            MatchCollection matches = Regex.Matches(band, pattern);

            // 2. ����Ƿ��ҵ���������������
            if (matches.Count >= 2)
            {
                // 3. ��ȡǰ����ƥ�䵽������
                if (double.TryParse(matches[0].Value, out double num1) &&
                    double.TryParse(matches[1].Value, out double num2))
                {
                    // 4. ȷ����Сֵ�����ֵ (��Ϊ����˳��һ��ȷ��)
                    minValue = Math.Min(num1, num2);
                    maxValue = Math.Max(num1, num2);
                    return true;
                }
            }

            // ���û���ҵ��������֣�����ת��ʧ��
            return false;
        }
        public XYPoint FilterXYPointData(XYPoint sourcePoint, double min, double max)
        {
            // 1. ������飺ȷ��Դ����XArrys �� YArrys �������ҳ���ƥ��
            if (sourcePoint == null ||
                sourcePoint.XArrys == null ||
                sourcePoint.YArrys == null ||
                sourcePoint.XArrys.Length != sourcePoint.YArrys.Length)
            {
                return null; // ������Ч������ null ���׳��쳣
            }

            // 2. ʹ�� Zip ������ X ����� Y ���鰴������Գ��������� (X, Y)
            var pairedData = sourcePoint.XArrys.Zip(sourcePoint.YArrys, (x, y) => new { X = x, Y = y });

            // 3. ɸѡ��ֻ���� X ֵ�� [min, max] ��Χ�ڵĶ�
            var filteredPairs = pairedData
                .Where(pair => pair.X >= min && pair.X <= max)
                .ToList();

            // 4. �����µ� XYPoint ���󣬲���ֻ� X ����� Y ����
            return new XYPoint
            {
                // �µ� Size ���Ի���ɸѡ������������߱���ԭ�е� Size
                Size = filteredPairs.Count,
                XArrys = filteredPairs.Select(pair => pair.X).ToArray(),
                YArrys = filteredPairs.Select(pair => pair.Y).ToArray()
            };
        }
        private void LoadProperites()
        {
            List<string> list = new List<string>();
            list.Add("MML806");
            //foreach (var item in Properties.Settings.Default.Amplifier)
            //{
            //    Console.WriteLine(item);
            //    list.Add(item);
            //}
            var parentItem = new NewTreeViewItem
            {
                Content = "Amplifier",
                Visible = true,
            };
            var children = new ObservableCollection<NewTreeViewItem>(
                list.Select(s => new NewTreeViewItem { Content = s, Parent = parentItem })
            );
            // 3. ���ӽڵ㼯�ϸ������ڵ�
            parentItem.Children = children;

            // 4. ��ӵ�������
            vm.Records.Add(parentItem);
            //vm.Records.Add(new NewTreeViewItem
            //{
            //    Content = "Amplifier",
            //    Visible = true,
            //    // Children.Add(new NewTreeViewItem { Content = "�����ձ�" }),
            //    Children = children
            //    // {
            //    //      new NewTreeViewItem{ Content = "MML806" }
            //    //}
            //});       
            //vm    }.Records.Last().AddChildren(Properties.Settings.Default.amp.Cast<string>());
        }
        private void btn_Add_TreeItem_Click(object sender, RoutedEventArgs e)
        {
            var menu = sender as MenuItem;
            var item = menu?.DataContext as NewTreeViewItem;
            Console.WriteLine(123);
            if (item != null)
            {
                Console.WriteLine(item.Content); // ? ��ȫ����
                PopupTitleChange(item.Content);
                //item.Children.Add(new NewTreeViewItem { Content = "MML808" });
                vm.Record = item;
                vm.PopupVisible = true;
                // ��Ҳ���Ը�ֵ�� vm.Record�������Ҫ��
                //vm.Record = item;
            }
        }

        private void Popup_OK_Click(object sender, RoutedEventArgs e)
        {
            vm.Record.Children.Add(new NewTreeViewItem { Content = vm.PopupText });
            vm.PopupVisible = false;
        }

        private void Popup_Cancel_Click(object sender, RoutedEventArgs e)
        {
            vm.PopupVisible = false;
        }

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var selectedItem = e.NewValue as NewTreeViewItem; // ?? ��󶨵���������

            if (selectedItem != null)
            {
                try
                {
                    if (selectedItem.Parent != null)
                    {
                        Console.WriteLine($"�����ˣ�{selectedItem.Content} {selectedItem.Parent.Content}");
                        // �������������� ViewModel�������˵�����ʾ�����
                        vm.ContentTitle =Global.TaskModel.TaskName+"-" +selectedItem.Parent.Content + "-" + selectedItem.Content;
                        table.SetBasicParameterRow(selectedItem.Content);
                        welComeStackPanel.Visibility = Visibility.Hidden;
                        contentGrid.Visibility = Visibility.Visible;
                    }
                    else
                    {

                        //vm.ContentTitle = selectedItem.Content;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                }
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            Console.WriteLine("exist");
            //foreach (var item in vm.Records)
            //{
            //    //���� Amplifier�µ����е������ļ�
            //    if (item.Content == "Amplifier")
            //    {
            //        Properties.Settings.Default.Amplifier.Clear();
            //        foreach (var child in item.Children)
            //        {
            //            Properties.Settings.Default.Amplifier.Add(child.Content);
            //        }
            //    }
            //}
            //Properties.Settings.Default.Save(); // ������� Save() �Ż�д����̣�
            base.OnClosed(e);
        }

        private void PopupTitleChange(string title)
        {
            vm.PopupTitle = $"Please Add A Chip Serial Number of {title}";

        }

        private async void Btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //var saveFileDialog = new SaveFileDialog
                //{
                //    Filter = "PPT Files (*.pptx)|*.pptx|All Files (*.*)|*.*",
                //    DefaultExt = ".pptx",
                //    FileName = "demo.pptx", // Ĭ���ļ���
                //    Title = "Save PPT File"
                //};
                //// ��ʾ�Ի��򣨱����� UI �̵߳��ã���
                //bool? result = saveFileDialog.ShowDialog();

                //if (result == true)
                //{
                //    string selectedPdfPath = saveFileDialog.FileName;

                //    // ����������� selectedPdfPath ������ PDF
                //    // ���磺File.WriteAllBytes(selectedPdfPath, pdfBytes);
                //    // �������� PDF �����߼�

                //    // ע�⣺��Ҫʹ��ԭ���� Path.Combine(...) Ӳ����·����
                //    // pdfFile = selectedPdfPath;
                //    await Task.Run(() => PPTChange(@"resources\files\T_MML806_V3.pptx", selectedPdfPath));

                //    string appDir = AppDomain.CurrentDomain.BaseDirectory;
                //    string pptFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pptx");
                //    // ���������ļ��Ի���


                //    string pdfFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pdf");

                //    // �ں�̨�߳�ִ�У����� UI ���ᣩ
                //    await Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));

                //    var pdfShown = new PdfShowWin();
                //    pdfShown.ShowPdf(pdfFile);
                //    pdfShown.Show();
                //}
                ////var pptPath = @"F:\PROJECT\ChipManualGeneration\exe\T_MML806_V3.pptx";
                string Data_Status_Str = "Are you sure you have finished preparing the data?";
                string File_Status_Str = $"Do you want to generate a PPT file for {Global.TaskModel.TaskName}?";
                string Check_Status_Str = $"Are you sure you have finished checking the {Global.TaskModel.TaskName}? ";
                const string tips = "Warning";
                switch ((UserPriority)Global.User.priority)
                { 
                    case UserPriority.Admin:
                        { 
                            
                        }
                        break;

                    case UserPriority.DataProvider:
                        {
                            var resoult = MessageBox.Show(Data_Status_Str, tips, MessageBoxButton.OK, MessageBoxImage.Warning);
                            if (resoult == MessageBoxResult.OK)
                            {
                                
                                var condition = filter.GetFileterCondition();

                                var sqlServer = new TaskRepository();
                                string cond = "";
                                foreach (var item in condition.FreqBands)
                                { 
                                    cond += item + ";";
                                }
                                 cond +=  "," + condition.Min + ',' + condition.Max + ',';
                                foreach (var item in condition.FreqBands)
                                {
                                    cond += item + ",";
                                }
                                var existingModel = await sqlServer.GetOperationByIdAsync(Convert.ToInt32(Global.TaskModel.ID));
                                if (existingModel == null)
                                {
                                    // ��¼�����ڣ�������Ҫ���в�������������׳��쳣
                                    // throw new Exception($"TaskID {taskId} not found for update.");
                                    return;
                                }
                                //var opModel = new OperationModel
                                //{
                                //    TaskID = Convert.ToInt32(Global.TaskModel.ID),
                                //    TaskName = Global.TaskModel.TaskName,
                                //    TimeStamp = DateTime.Now,
                                //    StartDateTime = condition.StartDateTime,
                                //    EndDateTime = condition.StopDateTime,
                                //    PN = condition.PN,
                                //    SN = condition.ON,
                                //    DataReady = true,
                                //    Condition = cond,
                                //};
                                existingModel.TimeStamp = DateTime.Now; // ����ʱ���
                                existingModel.StartDateTime = condition.StartDateTime;
                                existingModel.EndDateTime = condition.StopDateTime;
                                existingModel.PN = condition.PN;
                                existingModel.SN = condition.ON;
                                existingModel.DataReady = true;
                                existingModel.Condition = cond;
                                await sqlServer.UpdateOperationAsync(existingModel);

                                MessageBox.Show("Finished preparing the data,.\n Click the Finish button ", tips, MessageBoxButton.OK, MessageBoxImage.Warning);

                            }
                        }
                        break;

                    case UserPriority.PptMaker:
                        {
                            try
                            {
                                var resoult = MessageBox.Show(File_Status_Str, tips, MessageBoxButton.OK, MessageBoxImage.Warning);
                                if (resoult == MessageBoxResult.OK)
                                {
                                    string[] tmp = vm.ContentTitle.Split('-');
                                    string fileName = tmp.Last() + ".pptx";
                                    //string fileName =  "MML806.pptx";
                                    var saveFileDialog = new SaveFileDialog
                                    {
                                        Filter = "PPT Files (*.pptx)|*.pptx|All Files (*.*)|*.*",
                                        DefaultExt = ".pptx",
                                        FileName = fileName, // Ĭ���ļ���
                                        Title = "Save PPT File"
                                    };
                                    // ��ʾ�Ի��򣨱����� UI �̵߳��ã���
                                    bool? result1 = saveFileDialog.ShowDialog();

                                    if (result1 == true)
                                    {
                                        string finalSavePath = saveFileDialog.FileName;
                                        var sqlServer = new TaskRepository();
                                        var opModel = new OperationModel
                                        {
                                            TaskID = Convert.ToInt32(Global.TaskModel.ID),
                                            TaskName = Global.TaskModel.TaskName,
                                            FileReady = true,
                                            PptPath = finalSavePath,

                                        };
                                        await sqlServer.UpdateOperationAsync(opModel);
                                        try
                                        {
                                            string appDir = AppDomain.CurrentDomain.BaseDirectory;
                                            string picsPath = System.IO.Path.Combine(appDir, "resources", "pic");
                                            System.IO.Directory.Delete(picsPath, true);
                                            System.IO.Directory.CreateDirectory(picsPath);
                                            curves.SaveAllPlot(picsPath);
                                            PptDataModeFactory();


                                            string pptFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pptx");
                                            string pdfFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pdf");
                                            await Task.Run(() => GeneratePPT(pptFile));
                                            await Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));
                                            var pdfShown = new PdfShowWin();
                                            //pdfShown.Status = true;
                                            //PdfShowWin.PPTPath = pptFile;
                                            //PdfShowWin.PdfPath = pdfFile;
                                            pdfShown.ShowPdf(pdfFile);
                                            pdfShown.Show();
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
                                        }

                                        MessageBox.Show($"Generate a PPT file for {Global.TaskModel.TaskName}.\n Click the Finish button ", tips, MessageBoxButton.OK, MessageBoxImage.Warning);

                                    }
                                }


                            }
                            catch (Exception ex) {
                                MessageBox.Show(ex.Message, "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                        break;
                    case UserPriority.Reviewer:
                        {
                            var resoult = MessageBox.Show(Check_Status_Str, tips, MessageBoxButton.OK, MessageBoxImage.Warning);
                            if (resoult == MessageBoxResult.OK)
                            {
                                MessageBox.Show($"finished checking the { Global.TaskModel.TaskName}.\n Click the Finish button ", tips, MessageBoxButton.OK, MessageBoxImage.Warning);

                            }
                        }
                        break;


                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private bool PPTChange(string filePath = @"resources\files\T_MML806_V3.pptx", string tagetFilePath = @"resources\files\demo.pptx")
        {
            //filePath = @"F:\PROJECT\ChipManualGeneration\exe\T_MML806_V3.pptx";
            tagetFilePath = @"F:\PROJECT\ChipManualGeneration\exe\ChipManualGenerationSogt\ChipManualGenerationSogt\bin\Debug\resources\files\demo.pptx";
            bool success = false;
            try
            {

                if (!File.Exists(filePath))
                    throw new FileNotFoundException("δ�ҵ�PPT�ļ�", filePath);
                if (File.Exists(tagetFilePath))
                {
                    try
                    {
                        File.Delete(tagetFilePath);
                    }
                    catch (IOException ex)
                    {
                        throw new IOException($"Canno't delete file��{tagetFilePath}", ex);
                    }
                }
                File.Copy(filePath, tagetFilePath, overwrite: true);
                using (var presentationDoc = PresentationDocument.Open(tagetFilePath, isEditable: true))
                {
                    List<string> basicTableInfo = table.GetBasicTableInfo();
                    //pptDataModel.SliderMaster = new SliderMasterModel
                    //{
                    //    TopPN = basicTableInfo.ElementAt(0),
                    //    Version = basicTableInfo.ElementAt(1),
                    //    ProductName = basicTableInfo.ElementAt(2),
                    //    FrequencyRange = basicTableInfo.ElementAt(3),
                    //    RPN = basicTableInfo.ElementAt(4),
                    //    RightBarInfo = basicTableInfo.ElementAt(5)
                    //};
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Top PN}", "MML806");
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Version}", "V3.0.0");
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Product Name}", "GaAs MMIC Low Noise Amplifier");
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Frequency Range}", "45-90GHz");
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{R PN}", "MML806");
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{right bar info}", "GaAs Low Noise Amplifier MMIC 45 �C 90GHz");
                    var presentationPart = presentationDoc.PresentationPart;

                    // ��ȡ��һ�� slide part��ͨ����ϵ��
                    var slideId = presentationPart.Presentation.SlideIdList?.GetFirstChild<P.SlideId>();
                    //var slideId = presentationPart.Presentation.SlideIdList?.FirstOrDefault();
                    if (slideId == null)
                        throw new InvalidOperationException("PPT��û�лõ�Ƭ��");


                    // ? �ؼ���v3.3.0 ���� GetPartById ��ȡ SlidePart�������� OpenXmlPart������תΪ Slide��
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    var slide = slidePart.Slide;

                    var currentY = 1214000L; // ��ʼ Y λ��
                    const long verticalSpacing = 500000; // ��� 100,000 EMU

                    string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ��2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50��\nChip Size: 1.766 x 2.0 x 0.05mm";
                    string x = "Feautures\n" + table.GetFeatureTableInfo();
                    // ����ı��򲢻�ȡʵ�ʸ߶�
                    long height1 = PptModifier.AddTextBox(slide, x.TrimEnd(), 914400, currentY);
                    currentY += height1 + verticalSpacing + 100000;

                    string typicalApp = "Test Instrumentation\nMicrowave Radio & VSAT\nMilitary & Space\nTelecom Infrastructure\nFiber Optics";
                    long height2 = PptModifier.AddTextBox(slide, typicalApp, 914400, currentY);
                    currentY += height2 + verticalSpacing;

                    string info2 = "Electrical Specifications";
                    long height3 = PptModifier.AddTextBox(slide, info2, 914400, currentY);
                    currentY += height3 + 100000;

                    string info3 = "TA = +25\u2103, VD = +4V , VG=-0.4V , IDD = 119mA Typical";
                    long height4 = PptModifier.AddTextBox(slide, info3, 914400, currentY);
                    currentY += height4 + 10000;
                    string[,] tableData = new string[,]
                            {
                            { "Parameters", "Min.", "Typ.", "Max.", "Min.", "Typ.", "Max.", "Unit" },
                            { "Frequency",  "", "45-60", "", "", "60-90", "", "GHz" },
                            { "Small Signal Gain",  "14", "14.5", "", "16", "18", "", "dB" },
                            { "Gain Flatness",  "", "��1.0", "", "", "��1.0", "", "dB" },
                            { "Noise Figure",  "", "��1.0", "", "", "��1.0", "", "dB" },
                            { "P1dB - Output 1dB Compression",  "", "12", "", "", "14", "", "dBm" },
                            { "Psat - Saturated Output Power",  "", "12", "", "", "14", "", "dBm" },
                            { "OIP3 - Output Third Order Intercept",  "", "12", "", "", "14", "", "dBm" },
                            { "Input Return Loss",  "", "12", "", "", "14", "", "dB" },
                            { "Output Return Loss",  "", "12", "", "", "14", "", "dB" }
                    };

                    PptModifier.AddTable(slide, table.GetParameterTableInfo(), 914400, currentY, 6000000, 3800000);
                    currentY += 2000000 + verticalSpacing; // ���߶� + ���


                    // ����»õ�Ƭ
                    var newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    //AddNewSlideFromLayout

                    var newSlide = newSlidePart.Slide;
                    // ʾ�������»õ�Ƭ����ӱ��
                    int originX = 914400;

                    int originY = 1314000;
                    currentY = originY;
                    string info = "Measurement Plots: S-parameters\n TA = +25\u2103";// \u2103 �����϶ȵķ���
                    long height = PptModifier.AddTextBoxCenter(newSlide, info, originX, originY);
                    currentY += height + 50000;

                    var offsetX = 914400 + 2_500_000 + 700_000;
                    string pic1 = @"pic\0.png";
                    PptModifier.AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);

                    pic1 = @"pic\1.png";
                    PptModifier.AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);

                    pic1 = @"pic\2.png";
                    currentY += 2_000_000 + 300_000;
                    PptModifier.AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);
                    pic1 = @"pic\3.png";
                    PptModifier.AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);

                    currentY += 2_000_000 + 250_000;
                    info = "Measurement Plots: S-parameters\nVD=4.0V,VG=-0.5V";
                    height = PptModifier.AddTextBoxCenter(newSlide, info, originX, currentY);
                    currentY += height + 50000;
                    pic1 = @"pic\4.png";
                    PptModifier.AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);
                    pic1 = @"pic\5.png";
                    PptModifier.AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);





                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    info = "Absolute Maximum Ratings";// \u2103 �����϶ȵķ���
                    height = PptModifier.AddTextBoxUnderline(newSlide, info, originX, originY, 3500_000, 300000);

                    info = "Typical Supply Current vs. VD,VG";// \u2103 �����϶ȵķ���
                    height = PptModifier.AddTextBoxUnderline(newSlide, info, 914400 + 2700000 + 700000, originY, 3500_000, 300000);
                    currentY += height + 50000;
                    tableData = new string[,]
                            {

                            { "Drain Bias Voltage (VD)",  "+4.5V" },
                            { "Gate Bias Voltage (VG)",  "-2V to 0V" },
                            { "RF Input Power (RFIN)@(+4V)",  "+15"},
                            { "Channel Temperature",   "��1.0" },
                            { "Continuous Pdiss (T = 85 ��C)\n(derate 6.1mW/��C above 85 ��C)",  "12", },
                            { "Thermal Resistance\n (channel to die bottom)",   "12" },
                            { "Operating Temperature",  "55��C to +85 ��C"},
                            { "Storage Temperature",  "65��C to +150 ��C" },

                    };
                    PptModifier.AddTable2(newSlide, tableData, 914400, currentY, 2700000, 2500000);


                    tableData = new string[,]
                            {
                            { "VD (V)", "VG (V)", "IDD (mA)" },
                            { "+3.5",  "-0.38","118" },
                            { "+4.0",  "-0.40","119" },
                            { "+4.0",  "-0.50","71" },
                    };
                    PptModifier.AddTable(newSlide, tableData, 914400 + 2700000 + 700000, currentY, 2000000, 1500000);
                    pic1 = @"F:\PROJECT\ChipManualGeneration\exe\2.png";
                    PptModifier.AddImage(newSlidePart, pic1, 914400 + 2700000 + 600000, currentY + 220_0000, 50_0000, 50_0000);
                    info = "ELECTROSTATIC SENSITIVE DEVICE\n OBSERVE HANDLING PRECAUTIONS";
                    PptModifier.AddTextBox2(newSlide, info, 914400 + 3000000 + 600000 + 200_000, currentY + 220_0000);


                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    info = "Outline Drawing: \nAll Dimensions in ��m";
                    height = PptModifier.AddTextBoxCenter2(newSlide, info, originX, currentY);

                    pic1 = @"F:\PROJECT\ChipManualGeneration\exe\3.png";
                    PptModifier.AddImage(newSlidePart, pic1, 1014400, currentY + 800_000, 500_0000, 500_0000);

                    info = "Notes:\n1. Die thickness: 50��m\n2. VD bond pad is 75*75��m\u00B2 \n3. VG bond pad is 75*75��m\u00B2 \n4. RF IN/OUT bond pad is 50*86��m\u00B2 \n5. Bond pad metalization: Gold\n6. Backside metalization: Gold\n";
                    PptModifier.AddTextBox2(newSlide, info, 914400, currentY + 800_000 + 500_0000 + 100_000);

                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    info = "Assembly Drawing";
                    height = PptModifier.AddTextBoxCenter(newSlide, info, originX, currentY);

                    pic1 = @"F:\PROJECT\ChipManualGeneration\exe\4.png";
                    PptModifier.AddImage(newSlidePart, pic1, 1014400, currentY + height + 100, 550_0000, 350_0000);
                    tableData = new string[,]
                            {
                            { "Item", "Description" },
                            { "C1",  "100pF Example: Skyworks Part: SC10002430", },
                            { "C2",  "0.01?F Example: TDK\nPart:C1005X7R1H103K050BB (0402)", },
                            { "C3",  "0.1?F Example: TDK\nPart:C1005X7R1H104K050BB (0402)" },
                            { "R1",  "10? Example: Yageo\nPart:SR0402FR-7T10RL" },
                    };
                    PptModifier.AddTable3(newSlide, tableData, 914400, currentY + height + 100 + 350_0000 + 50_000, 350_0000, 200_0000);

                    tableData = new string[,]
                            {
                            { "No", "Function","Description" },
                            { "1",  "RF IN", "RF signal input terminal; no blocking capacitor required. "},
                            { "2",  "RF OUT","RF signal output terminal; no blocking capacitor required." },
                            { "3",  "VD","Drain Biases for the Amplifier ; An external biasing circuit is required." },
                            { "4",  "VG", "Gate Biases for the Amplifier ; An external biasing circuit is required."},
                            { "5",  "Die Bottom", "Die bottom must be connected to RF and dc ground."},
                    };
                    PptModifier.AddTable3(newSlide, tableData, 914400, currentY + height + 100 + 350_0000 + 1000 + 200_0000 + 150_000, 600_0000, 200_0000);


                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;

                    pic1 = @"F:\PROJECT\ChipManualGeneration\exe\5.png";
                    PptModifier.AddImage(newSlidePart, pic1, 2114400, currentY, 300_0000, 300_0000);
                    info = "Biasing and Operation";
                    height = PptModifier.AddTextBoxCenter2(newSlide, info, originX, currentY + 300_0000 + 100_0000);

                    info = "Turn ON procedure: \n1.  Connect GND to RF and dc ground.\n2.  Set the gate bias voltages,VG to ?2V.\n3.  Set the drain bias voltages VD to +4V .\n4.  Increase the gate bias voltages to achieve a quiescent supply current of 82 mA.\n5.  Apply RF signal.?\n \nTurn OFF procedure: \n1.  Turn off the RF signal.\n2.  Decrease the gate bias voltages, VG to ?2V to achieve a IDQ = 0 mA (approximately).\n3.  Decrease the drain bias voltages to 0 V.\n4.  Increase the all gate bias voltages to 0 V.\n";
                    PptModifier.AddTextBox2(newSlide, info, 914400, currentY + 300_0000 + 100_0000 + height + 800);

                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    info = "Mounting & Bonding Technigues for MMICs";
                    height = PptModifier.AddTextBoxCenter(newSlide, info, originX, currentY);
                    pic1 = @"F:\PROJECT\ChipManualGeneration\exe\6.png";
                    PptModifier.AddImage(newSlidePart, pic1, 1514400, currentY, 500_0000, 150_0000);
                    info = "Direct Mounting\n1.Typically, the die is mounted directly on the ground plane.\n2.If the thickness difference between the?substrate (thickness?c)?and the?die (thickness?d)?exceeds?0.05 mm (i.e.,?c?�C?d?> 0.05 mm), it is recommended to first mount the die on a?heat spreader, then attach the heat spreader to the ground plane.\r\n3.Heat Spreader Material: Molybdenum-copper (MoCu) alloy is commonly used.\r\n4.Heat Sink Thickness (b): Should be within the range of?(c?�C?d?�C 0.05 mm)?to?(c?�C?d?+ 0.05 mm).\r\n5.Spacing (a): The gap between the bare die and the 50�� transmission line should typically be?0.05 mm to 0.1 mm.\r\nIf the application frequency is higher than 40GHz, then this gap is recommended to be 0.05mm\r\nWire Bonding Interconnection\r\nThe connection between the die and the 50�� transmission line is usually made using?25 ?m diameter gold (Au) wires, bonded via?wedge bonding?or?ball bonding?processes.\r\nDie Attachment Methods\r\n1.Conductive Epoxy:\r\nAfter adhesive application, cure according to the manufacturer��s recommended temperature profile.\r\n2.Au-Sn80/20 Eutectic Bonding:\r\nUse preformed?Au-Sn80/20 solder preforms.\r\nPerform bonding in an inert atmosphere (N? or forming gas: 90% N? + 10% H?).\r\nKeep the time above?320��C?to?less than 20 seconds?to prevent excessive intermetallic formation.\r\n";
                    string infox = info.Replace("\r\n", "\n");
                    height = PptModifier.AddTextBox5(newSlide, infox, 914400, currentY + 150_0000 + 1_0000);
                    info = "Miller MMIC Inc. All rights reserved\r\nMiller MMIC, Inc. holds exclusive rights to the information presented in its Data Sheet and any accompanying materials. As a premier supplier of cutting-edge RF solutions, Miller MMIC has made this information easily accessible to its clients.\r\nAlthough Miller MMIC believes the information provided in its Data Sheet to be trustworthy, the company does not offer any guarantees as to its accuracy. Therefore, Miller MMIC bears no responsibility for the use of this information. It is worth mentioning that the information within the Data Sheet may be altered without prior notification.\r\nCustomers are encouraged to obtain and verify the most recent and pertinent information before placing any orders for Miller MMIC products. The information in the Data Sheet does not confer, either explicitly or implicitly, any rights or licenses with regards to patents or other forms of intellectual property to any third party.\r\nThe information provided in the Data Sheet, or its utilization, does not bestow any patent rights, licenses, or other forms of intellectual property rights to any individual or entity, whether in regards to the information itself or anything described by such information. Furthermore, Miller MMIC products are not intended for use as critical components in applications where failure could result in severe injury or death, such as medical or life-saving equipment, or life-sustaining applications, or in any situation where failure could cause serious personal injury or death.\r\n";
                    infox = info.Replace("\r\n", "\n");
                    PptModifier.AddTextBox4(newSlide, infox, 914400, currentY + height + 350_0000);

                    slide.Save();
                    success = true;
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                success = false;
            }
            return success;
        }

        public void GeneratePPT(string tagetFilePath)
        {

            tagetFilePath = @"F:\PROJECT\ChipManualGeneration\exe\ChipManualGenerationSogt\ChipManualGenerationSogt\bin\Debug\resources\files\demo.pptx";
            string filePath = @"resources\files\T_MML806_V3.pptx";
            bool success = false;
            try
            {

                if (!File.Exists(filePath))
                    throw new FileNotFoundException("δ�ҵ�PPT�ļ�", filePath);
                if (File.Exists(tagetFilePath))
                {
                    try
                    {
                        File.Delete(tagetFilePath);
                    }
                    catch (IOException ex)
                    {
                        throw new IOException($"Canno't delete file��{tagetFilePath}", ex);
                    }
                }
                File.Copy(filePath, tagetFilePath, overwrite: true);
                using (var presentationDoc = PresentationDocument.Open(tagetFilePath, isEditable: true))
                {
                    //�޸�ĸ�����Ϣ
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Top PN}", pptDataModel.SliderMaster.TopPN);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Version}", pptDataModel.SliderMaster.Version);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Product Name}", pptDataModel.SliderMaster.ProductName);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Frequency Range}", pptDataModel.SliderMaster.FrequencyRange);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{R PN}", pptDataModel.SliderMaster.TopPN);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{right bar info}", pptDataModel.SliderMaster.RightBarInfo);

                    var presentationPart = presentationDoc.PresentationPart;

                    // ��ȡ��һ�� slide part��ͨ����ϵ��
                    var slideId = presentationPart.Presentation.SlideIdList?.GetFirstChild<P.SlideId>();
                    //var slideId = presentationPart.Presentation.SlideIdList?.FirstOrDefault();
                    if (slideId == null)
                        throw new InvalidOperationException("PPT��û�лõ�Ƭ��");

                    #region ��һҳ
                    // ? �ؼ���v3.3.0 ���� GetPartById ��ȡ SlidePart�������� OpenXmlPart������תΪ Slide��
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    var slide = slidePart.Slide;

                    var currentY = 1214000L; // ��ʼ Y λ��
                    const long verticalSpacing = 500000; // ��� 100,000 EMU

                    //string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ��2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50��\nChip Size: 1.766 x 2.0 x 0.05mm";
                    //string x = "Feautures\n" + table.GetFeatureTableInfo();

                    // д��feature
                    long height1 = PptModifier.AddTextBox(slide, pptDataModel.FirstPage.FeaturesText.TrimEnd(), 914400, currentY);
                    // ����ı��򲢻�ȡʵ�ʸ߶�
                    currentY += height1 + verticalSpacing + 100000;

                    long height2 = PptModifier.AddTextBox(slide, pptDataModel.FirstPage.TypicalApplicationsText, 914400, currentY);
                    currentY += height2 + verticalSpacing;

                    string info2 = "Electrical Specifications";
                    long height3 = PptModifier.AddTextBox(slide, pptDataModel.FirstPage.ElectricalSpecsTitle, 914400, currentY);
                    currentY += height3 + 100000;

                    string info3 = "TA = +25\u2103, VD = +4V , VG=-0.4V , IDD = 119mA Typical";
                    long height4 = PptModifier.AddTextBox(slide, pptDataModel.FirstPage.ElectricalSpecsCondition, 914400, currentY);
                    currentY += height4 + 10000;
                    //PptModifier.AddTable(slide, table.GetParameterTableInfo(), 914400, currentY, 6000000, 3800000);
                    PptModifier.AddTable6(slide, pptDataModel.FirstPage.ParameterTableData, 914400, currentY, 6000000, 380_0000);
                    currentY += 2000000 + verticalSpacing; // ���߶� + ���

                    PptModifier.AddTextBox(slide, pptDataModel.FirstPage.FunctionalBlockDiagramImage.ImageName.TrimEnd(), 914400 + 3_250_000,1314000 + 150_000);
                    PptModifier.AddImage(slidePart, pptDataModel.FirstPage.FunctionalBlockDiagramImage.ImagePath, 914400+ 3200_000, 1314000+ 400_000, 240_0000, 270_0000);


                    #endregion

                    // ������4�ı����� ������ǣ� �ͱ�ʾ������

                    #region �������ҳ
                    int curvePageCount;

                    if (pptDataModel.CurvesImagePage.CurveImagesPath.Count % 6 == 0)
                    {
                        // ҳ��Ϊ6����0����
                        curvePageCount = pptDataModel.CurvesImagePage.CurveImagesPath.Count / 6;
                    }
                    else
                    {
                        curvePageCount = pptDataModel.CurvesImagePage.CurveImagesPath.Count / 6 + 1;
                    }
                    int originX = 914400;
                    var offsetX = originX + 2_500_000 + 700_000;

                    

                    int textOriginY = 1314000;
                    int textOriginY2 = textOriginY + 2_000_000 + 600_000;
                    int textOriginY3 = textOriginY2 + 2_000_000 + 600_000;

                    int originY = textOriginY + 500_000;//��һ��
                    int originY2 = textOriginY2 + 500_000;//�ڶ���
                    int originY3 = textOriginY3 + 500_000;//������

                    List<(int x, int y)> imgagePositions = new List<(int x, int y)>();// ��λλ�ã� ͼ��λ�ó���Ϊ 3��2�еı�
                    imgagePositions.Add((originX, originY));
                    imgagePositions.Add((offsetX, originY));

                    imgagePositions.Add((originX, originY2));
                    imgagePositions.Add((offsetX, originY2));

                    imgagePositions.Add((originX, originY3));
                    imgagePositions.Add((offsetX, originY3));


                    List<(int x, int y)> textBoxPositions = new List<(int x, int y)>();// ��λλ�ã� ͼ��λ�ó���Ϊ 3��2�еı�
                    textBoxPositions.Add((originX, textOriginY));
                    textBoxPositions.Add((offsetX, textOriginY));

                    textBoxPositions.Add((originX, textOriginY2));
                    textBoxPositions.Add((offsetX, textOriginY2));

                    textBoxPositions.Add((originX, textOriginY3));
                    textBoxPositions.Add((offsetX, textOriginY3));

                    //int textIndex = 0;
                    //int index = 0;
                    //for (int i = 0; i < curvePageCount; i++)
                    //{

                    //    var newSlidePart1 = PptModifier.AddNewSlideFromLayout(presentationPart);
                    //    //AddNewSlideFromLayout

                    //    var newSlide1 = newSlidePart1.Slide;
                    //    // ʾ�������»õ�Ƭ����ӱ��

                    //    currentY = originY;

                    //    long height5 = 0;
                    //    for (; index < pptDataModel.CurvesImagePage.CurveImagesPath.Count; )
                    //    {
                    //        string path = pptDataModel.CurvesImagePage.CurveImagesPath.ElementAt(index);
                    //        Console.WriteLine(path);
                    //        PptModifier.AddImage(newSlidePart1, path, imgagePositions.ElementAt(index % 6).x, imgagePositions.ElementAt(index % 6).y, 2_500_000, 2_000_000);


                    //        string text = pptDataModel.CurvesImagePage.CurveTitles.ElementAt(textIndex);
                    //        if (index==0)
                    //        {
                    //            height5 = PptModifier.AddTextBoxCenter(newSlide1, text, textBoxPositions.ElementAt(0).x, textBoxPositions.ElementAt(0).y);
                    //            textIndex++;
                    //        }
                    //        else
                    //        {
                    //            if (index < 4)
                    //                continue;
                    //            // ��ȥ��һ������
                    //            if (((index - 4) / 4) % 2 == 1) // 4����������  ��������ʾ���� s������ ż������ʾ����nf�� psat�� oip3�� pxdb
                    //            {
                    //                if (index % 4 == 0)//4��s������ ֻ��Ҫ��һ�α���
                    //                {
                    //                    PptModifier.AddTextBoxCenter(newSlide1, text, textBoxPositions.ElementAt(index % 6).x, textBoxPositions.ElementAt(index % 6).y);
                    //                    textIndex++;
                    //                }
                    //            }

                    //            else
                    //            {// ż��������ʾ���� nf�� psat�� oip3�� pxdb
                    //             //4��s������ ֻ��Ҫ��һ�α���
                    //                PptModifier.AddTextBoxCenterWH(newSlide1, text, textBoxPositions.ElementAt(index % 6).x, textBoxPositions.ElementAt(index % 6).y, 350_0000, 350_0000);
                    //                textIndex++;
                    //            }



                    //        }

                    //        if (index > 0 && index % 6 == 0)
                    //        {
                    //            index++;
                    //            break;// 6��һ�飬 ����ֻ��Ҫѭ��6�Σ� �����µ�һҳ
                    //        }

                    //    }

                    //}
                   int IMAGES_PER_SLIDE = 6;
                   int IMAGE_TITLE_SKIP_GROUP = 4;

                    {
                        int imageIndex = 0; // ���ڸ����ܵ�ͼƬ����
                        int titleIndex = 0; // ���ڸ����ܵı�������
                        int imagesCount = pptDataModel.CurvesImagePage.CurveImagesPath.Count;
                        for (int i = 0; i < curvePageCount; i++)
                        {
                            var newSlidePart1 = PptModifier.AddNewSlideFromLayout(presentationPart);
                            var newSlide1 = newSlidePart1.Slide;

                            for (int j = 0; j < IMAGES_PER_SLIDE; j++)
                            {
                                // ����Ƿ��ѴﵽͼƬ����ĩβ
                                if (imageIndex >= pptDataModel.CurvesImagePage.CurveImagesPath.Count)
                                {
                                    break;
                                }

                                // --- ͼƬ���� ---
                                string imagePath = pptDataModel.CurvesImagePage.CurveImagesPath.ElementAt(imageIndex);
                                Console.WriteLine(imagePath);

                                PptModifier.AddImage(
                                    newSlidePart1,
                                    imagePath,
                                    imgagePositions.ElementAt(j).x,
                                    imgagePositions.ElementAt(j).y,
                                    2_500_000, 2_000_000
                                );

                                // --- ���⴦��ֲ�������ʼ�� ---
                                string titleText = "";
                                bool shouldAddTitle = false;

                                // Ĭ��ֵ���Է�ֹδ���� else ��ʱʹ��
                                int adjustedIndex = -1;
                                bool isSParameterGroup = false;

                                // �����������Ƿ�Խ�磨������ ElementAt() ֮ǰ��
                                if (titleIndex >= pptDataModel.CurvesImagePage.CurveTitles.Count)
                                {
                                    imageIndex++;
                                    continue; // û�и��������
                                }
                                titleText = pptDataModel.CurvesImagePage.CurveTitles.ElementAt(titleIndex);

                                // --- ���ӱ�������߼� ---
                                if ((imagesCount/4) % 2 != 0) // ���ʱ���е�������4��ͼ  ͼƬ��ĿΪ 4 + 8* n  �� ��������4��ֻ����s������  ����8��һ���ʾ ������ nf, psat, oip3, pxdb
                                {
                                    if (imageIndex == 0)
                                    {
                                        // ��һ��ͼƬ��������ӱ���
                                        shouldAddTitle = true;
                                        // ע�⣺���� j Ҳ�� 0
                                    }
                                    else if (imageIndex < IMAGE_TITLE_SKIP_GROUP)
                                    {
                                        // imageIndex > 0 ��С�� 4 �������ԭʼ�߼��� continue (����ӱ���)
                                        shouldAddTitle = false;
                                    }
                                    else
                                    {
                                        // �������ڵ��� 4 �����
                                        adjustedIndex = imageIndex - IMAGE_TITLE_SKIP_GROUP;

                                        // ��������������Ƿ����� S ������ (4��������)
                                        // ԭʼ�߼�: ((index - 4) / 4) % 2 == 1
                                        isSParameterGroup = (adjustedIndex / IMAGE_TITLE_SKIP_GROUP) % 2 == 0;

                                        if (isSParameterGroup)
                                        {
                                            // S �����飺ÿ�� 4 ��ͼ��ֻ��Ҫ��һ�α���
                                            if (adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0)
                                            {
                                                shouldAddTitle = true;
                                            }
                                        }
                                        else
                                        {
                                            // �� S �����飺ÿ��ͼ���ӱ��� (ż����)
                                            shouldAddTitle = true;
                                        }
                                    }

                                }
                                else
                                {
                                    if (imageIndex == 0)
                                    {
                                        // ��һ��ͼƬ��������ӱ���
                                        shouldAddTitle = true;
                                        // ע�⣺���� j Ҳ�� 0
                                    }
                                    else if (imageIndex < IMAGE_TITLE_SKIP_GROUP)
                                    {
                                        // imageIndex > 0 ��С�� 4 �������ԭʼ�߼��� continue (����ӱ���)
                                        shouldAddTitle = false;
                                    }
                                    else
                                    {
                                        // �������ڵ��� 4 �����
                                        adjustedIndex = imageIndex - IMAGE_TITLE_SKIP_GROUP;

                                        // ��������������Ƿ����� S ������ (4��������)
                                        // ԭʼ�߼�: ((index - 4) / 4) % 2 == 1
                                        isSParameterGroup = (adjustedIndex / IMAGE_TITLE_SKIP_GROUP) % 2 == 1;

                                        if (isSParameterGroup)
                                        {
                                            // S �����飺ÿ�� 4 ��ͼ��ֻ��Ҫ��һ�α���
                                            if (adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0)
                                            {
                                                shouldAddTitle = true;
                                            }
                                        }
                                        else
                                        {
                                            // �� S �����飺ÿ��ͼ���ӱ��� (ż����)
                                            shouldAddTitle = true;
                                        }
                                    }

                                }



                                //if (imageIndex == 0)
                                //        {
                                //            // ��һ��ͼƬ��������ӱ���
                                //            shouldAddTitle = true;
                                //            // ע�⣺���� j Ҳ�� 0
                                //        }
                                //        else if (imageIndex < IMAGE_TITLE_SKIP_GROUP)
                                //        {
                                //            // imageIndex > 0 ��С�� 4 �������ԭʼ�߼��� continue (����ӱ���)
                                //            shouldAddTitle = false;
                                //        }
                                //        else
                                //        {
                                //            // �������ڵ��� 4 �����
                                //            adjustedIndex = imageIndex - IMAGE_TITLE_SKIP_GROUP;

                                //            // ��������������Ƿ����� S ������ (4��������)
                                //            // ԭʼ�߼�: ((index - 4) / 4) % 2 == 1
                                //            isSParameterGroup = (adjustedIndex / IMAGE_TITLE_SKIP_GROUP) % 2 == 1;

                                //            if (isSParameterGroup)
                                //            {
                                //                // S �����飺ÿ�� 4 ��ͼ��ֻ��Ҫ��һ�α���
                                //                if (adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0)
                                //                {
                                //                    shouldAddTitle = true;
                                //                }
                                //            }
                                //            else
                                //            {
                                //                // �� S �����飺ÿ��ͼ���ӱ��� (ż����)
                                //                shouldAddTitle = true;
                                //            }
                                //        }



                                // --- ִ�б������ ---
                                if (shouldAddTitle)
                                {
                                    // ����� if ������Ҫע�⣺ֻ�е� isSParameterGroup �� true 
                                    // �� adjustedIndex % 4 == 0 (�� S ������ĵ�һ��ͼ) ʱ����ʹ�� AddTextBoxCenter

                                    // imageIndex == 0 ����� isSParameterGroup �� adjustedIndex ����Ĭ��ֵ��
                                    // ��Ҫȷ���������� AddTextBoxCenter ������

                                    // �Ż���ʹ��һ������ȷ�����������ֵ��÷���
                                    bool useAddTextBoxCenter = (imageIndex == 0) ||
                                                               (isSParameterGroup && adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0);


                                    if (useAddTextBoxCenter)
                                    {
                                        // ���� imageIndex=0 ����� �� S ������ֻ���һ�ε����
                                        PptModifier.AddTextBoxCenter(
                                            newSlide1,
                                            titleText,
                                            textBoxPositions.ElementAt(j).x,
                                            textBoxPositions.ElementAt(j).y,
                                            1200
                                        );
                                    }
                                    else
                                    {
                                        // ���ڷ� S ����������
                                        PptModifier.AddTextBoxCenterWH(
                                            newSlide1,
                                            titleText,
                                            textBoxPositions.ElementAt(j).x,
                                            textBoxPositions.ElementAt(j).y,
                                            290_0000, 350_0000
                                        );
                                    }
                                    titleIndex++; // ֻ��������˱��������ӱ�������
                                }


                                imageIndex++; // �����Ƿ���ӱ��⣬ͼƬ������Ҫ����

                                // --- ԭʼ�ķ�ҳ�ж� ---
                                if (imageIndex > 0 && imageIndex % IMAGES_PER_SLIDE == 0)
                                {
                                    // �Ѿ������� 6 ��ͼ��׼��������һҳ
                                    break;
                                }
                            } // ������ǰ�õ�Ƭ�� 6 ��ͼƬѭ��
                        } // ������ҳѭ��
                    }
                    #endregion





                    #region ������5ҳ
                    var newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    var newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    //info = "Absolute Maximum Ratings";// \u2103 �����϶ȵķ���
                    PptModifier.AddTextBoxUnderline(newSlide, pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTableTitle, 814400, originY, 3500_000, 300000);

                    //info = "Typical Supply Current vs. VD,VG";// \u2103 �����϶ȵķ���
                    var height = PptModifier.AddTextBoxUnderline(newSlide, pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTableTitle, 914400 + 2700000 + 700000, originY, 3500_000, 300000);
                    currentY += height + 50000;
                    PptModifier.AddTable7(newSlide, pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTable, 814400, currentY, 330_0000, 3000000);
                    PptModifier.AddTableAverageWidth(newSlide, pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTable, 101_4400 + 2700000 + 700000, currentY, 2300000, 110_0000);
                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\2.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.EndToFront5Page.WarningImage.ImagePath, 914400 + 2700000 + 600000, currentY + 220_0000, 50_0000, 50_0000);
                    //info = "ELECTROSTATIC SENSITIVE DEVICE\n OBSERVE HANDLING PRECAUTIONS";
                    PptModifier.AddTextBox2(newSlide, pptDataModel.EndToFront5Page.WarningText, 91_4400 + 3000000 + 600000 + 200_000, currentY + 220_0000);

                    #endregion

                    #region ������4ҳ
                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    //info = "Outline Drawing: \nAll Dimensions in ��m";
                    //height = PptModifier.AddTextBoxCenter2(newSlide, pptDataModel.EndToFront4Page.PinImage.ImageName, originX, currentY);
                    height = PptModifier.AddTextBoxCenter2(newSlide, "Outline Drawing: \nAll Dimensions in ��m", originX, currentY);

                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\3.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.EndToFront4Page.PinImage.ImagePath, 1014400, currentY + 450_000, 500_0000, 500_0000);

                    //info = "Notes:\n1. Die thickness: 50��m\n2. VD bond pad is 75*75��m2 \n3. VG bond pad is 75*75��m2 \n4. RF IN/OUT bond pad is 50*86��m2 \n5. Bond pad metalization: Gold\n6. Backside metalization: Gold\n";
                    PptModifier.AddTextBox2(newSlide, pptDataModel.EndToFront4Page.NoteText, 914400, currentY + 800_000 + 500_0000 + 100_000);



                    #endregion


                    #region ������3ҳ
                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    //info = "Assembly Drawing";
                    height = PptModifier.AddTextBoxCenter(newSlide, pptDataModel.EndToFront3Page.StructImage.ImageName, originX, currentY);

                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\4.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.EndToFront3Page.StructImage.ImagePath, 1014400, currentY + height + 100, 550_0000, 350_0000);

                    PptModifier.AddTable4(newSlide, pptDataModel.EndToFront3Page.Description, 914400, currentY + height + 100 + 350_0000 + 50_000, 350_0000, 200_0000);

                    PptModifier.AddTable4(newSlide, pptDataModel.EndToFront3Page.Description2, 914400, currentY + height + 100 + 350_0000 + 1000 + 200_0000 + 150_000, 600_0000, 200_0000);

                    #endregion




                    #region  �����ڶ�ҳ
                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;

                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\5.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.EndToFront2Page.StructImage.ImagePath, 251_4400, currentY, 300_0000, 300_0000);
                    //info = "Biasing and Operation";
                    height = PptModifier.AddTextBoxCenter2(newSlide, pptDataModel.EndToFront2Page.Title, originX, currentY + 300_0000 + 100_0000);

                    //string info = "Turn ON procedure: \n1.  Connect GND to RF and dc ground.\n2.  Set the gate bias voltages,VG to ?2V.\n3.  Set the drain bias voltages VD to +4V .\n4.  Increase the gate bias voltages to achieve a quiescent supply current of 82 mA.\n5.  Apply RF signal.?\n \nTurn OFF procedure: \n1.  Turn off the RF signal.\n2.  Decrease the gate bias voltages, VG to ?2V to achieve a IDQ = 0 mA (approximately).\n3.  Decrease the drain bias voltages to 0 V.\n4.  Increase the all gate bias voltages to 0 V.\n";

                    string info = pptDataModel.EndToFront2Page.TurnOn + "\r\n" + pptDataModel.EndToFront2Page.TurnOff;

                    PptModifier.AddTextBox8(newSlide, info, 914400, currentY + 300_0000 + 100_0000 + height + 800);

                    #endregion
                    #region
                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    info = "Mounting & Bonding Technigues for MMICs";
                    height = PptModifier.AddTextBoxCenter(newSlide, pptDataModel.LastPage.Image.ImageName, originX, currentY);
                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\6.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.LastPage.Image.ImagePath, 1514400, currentY+40_0000, 500_0000, 150_0000);
                    info = "Direct Mounting\n1.Typically, the die is mounted directly on the ground plane.\n2.If the thickness difference between the?substrate (thickness?c)?and the?die (thickness?d)?exceeds?0.05 mm (i.e.,?c?�C?d?> 0.05 mm), it is recommended to first mount the die on a?heat spreader, then attach the heat spreader to the ground plane.\r\n3.Heat Spreader Material: Molybdenum-copper (MoCu) alloy is commonly used.\r\n4.Heat Sink Thickness (b): Should be within the range of?(c?�C?d?�C 0.05 mm)?to?(c?�C?d?+ 0.05 mm).\r\n5.Spacing (a): The gap between the bare die and the 50�� transmission line should typically be?0.05 mm to 0.1 mm.\r\nIf the application frequency is higher than 40GHz, then this gap is recommended to be 0.05mm\r\nWire Bonding Interconnection\r\nThe connection between the die and the 50�� transmission line is usually made using?25 ?m diameter gold (Au) wires, bonded via?wedge bonding?or?ball bonding?processes.\r\nDie Attachment Methods\r\n1.Conductive Epoxy:\r\nAfter adhesive application, cure according to the manufacturer��s recommended temperature profile.\r\n2.Au-Sn80/20 Eutectic Bonding:\r\nUse preformed?Au-Sn80/20 solder preforms.\r\nPerform bonding in an inert atmosphere (N? or forming gas: 90% N? + 10% H?).\r\nKeep the time above?320��C?to?less than 20 seconds?to prevent excessive intermetallic formation.\r\n";
                    string infox = info.Replace("\r\n", "\n");
                    height = PptModifier.AddTextBox5(newSlide, pptDataModel.LastPage.Text1, 914400, currentY + 150_0000 + 80_0000);
                    info = "Miller MMIC Inc. All rights reserved\r\nMiller MMIC, Inc. holds exclusive rights to the information presented in its Data Sheet and any accompanying materials. As a premier supplier of cutting-edge RF solutions, Miller MMIC has made this information easily accessible to its clients.\r\nAlthough Miller MMIC believes the information provided in its Data Sheet to be trustworthy, the company does not offer any guarantees as to its accuracy. Therefore, Miller MMIC bears no responsibility for the use of this information. It is worth mentioning that the information within the Data Sheet may be altered without prior notification.\r\nCustomers are encouraged to obtain and verify the most recent and pertinent information before placing any orders for Miller MMIC products. The information in the Data Sheet does not confer, either explicitly or implicitly, any rights or licenses with regards to patents or other forms of intellectual property to any third party.\r\nThe information provided in the Data Sheet, or its utilization, does not bestow any patent rights, licenses, or other forms of intellectual property rights to any individual or entity, whether in regards to the information itself or anything described by such information. Furthermore, Miller MMIC products are not intended for use as critical components in applications where failure could result in severe injury or death, such as medical or life-saving equipment, or life-sustaining applications, or in any situation where failure could cause serious personal injury or death.\r\n";
                    infox = info.Replace("\r\n", "\n");
                    PptModifier.AddTextBox4(newSlide, pptDataModel.LastPage.Text2, 914400, currentY + height + 350_0000);
                    #endregion
                    slide.Save();
                    success = true;
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                success = false;
            }


        }


        private async void Btn_Refresh_Clicked(object sender, RoutedEventArgs e)
        {
            test2();
        }

        private async void Btn_Preview_Clicked(object sender, RoutedEventArgs e)
        {

            try
            {
               
                //var pptPath = @"F:\PROJECT\ChipManualGeneration\exe\T_MML806_V3.pptx";
                await Task.Run(() => PPTChange());
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string pptFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pptx");
                string pdfFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pdf");

                //string pptFile = @"resources\files\demo.pptx"; �������ʽִ��ת����ת��ʧ��
                //string pdfFile = @"resources\files\demo.pdf";
                // �ں�̨�߳�ִ�У����� UI ���ᣩ
                await Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));

                var pdfShown = new PdfShowWin();
                pdfShown.Status = true;
                PdfShowWin.PPTPath = pptFile;
                PdfShowWin.PdfPath = pdfFile;
                pdfShown.ShowPdf(pdfFile);
                pdfShown.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
            }




        }

        private void HandleQueryFinished(object sender, EventArgs e)
        {
            Console.WriteLine("Logger: Query has finished.");
            
            ////��Application���ǡ�Microsoft.Office.Interop.PowerPoint.Application���͡�System.Windows.Application��֮��Ĳ���ȷ������
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                contentGridRight.Visibility = Visibility.Visible;
                test2();  // �����߳�ִ�� UI ����
            });

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {

                //GeneratePPTBtn.Visibility = Visibility.Visible;
                //PreviewPPTBtn.Visibility = Visibility.Visible;
                tableTab.Visibility = Visibility.Visible;
                curveTab.Visibility = Visibility.Visible;
                imgTab.Visibility = Visibility.Visible;
                //MessageBox.Show("Query has finished.", "Tips", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }

        private  async void Btn_Preview_PPT_Model_Clicked(object sender, RoutedEventArgs e)
        {
            //string pdfFile = "";
            //if (vm.ContentTitle == "Amplifier--MM809")
            //    pdfFile = @"F:\PROJECT\ChipManualGeneration\�Ŵ���\MML806_V3.pdf";
            //else if (vm.ContentTitle == "Amplifier--MM808")
            //    pdfFile = "F:\\PROJECT\\ChipManualGeneration\\�Ŵ���\\MML814_V3.0.1.pdf";
            //try
            //{
            //    var pdfShown = new PdfShowWin();
            //    pdfShown.ShowPdf(pdfFile);
            //    pdfShown.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
            //}
            try
            {
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string picsPath = System.IO.Path.Combine(appDir, "resources", "pic");
                System.IO.Directory.Delete(picsPath, true);
                System.IO.Directory.CreateDirectory(picsPath);
                curves.SaveAllPlot(picsPath);
                PptDataModeFactory();
                
                
                string pptFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pptx");
                string pdfFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pdf");
                await Task.Run(() => GeneratePPT(pptFile));
                await Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));
                var pdfShown = new PdfShowWin();
                //pdfShown.Status = true;
                //PdfShowWin.PPTPath = pptFile;
                //PdfShowWin.PdfPath = pdfFile;
                pdfShown.ShowPdf(pdfFile);
                pdfShown.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }


        private void HandleHomeBtnClick(object sender, EventArgs e)
        {
            HiddenAll();
            home.Visibility = Visibility.Visible;
        }
        private void HandleOperationBtnClick(object sender, EventArgs e)
        {
            HiddenAll();
            if (Global.OperationModel != null)
            {
                if (Global.OperationModel.DataReady)
                {
                    treeView.Visibility = Visibility.Visible;
                    separator.Visibility = Visibility.Visible;
                    contentGrid.Visibility = Visibility.Visible;
                    return;
                }
            }
            if (Global.TaskModel != null && vm.ContentTitle != null)
            {
                treeView.Visibility = Visibility.Visible;
                separator.Visibility = Visibility.Visible;
                contentGrid.Visibility = Visibility.Visible;
            }
            else if (Global.TaskModel != null && vm.ContentTitle == null)
            {

                treeView.Visibility = Visibility.Visible;
                welComeStackPanel.Visibility = Visibility.Visible;
            }
            else
            {
                //treeView.Visibility = Visibility.Visible;
                welComeStackPanel.Visibility = Visibility.Visible;
            }

                
        }
        private void HandleLogBtnClick(object sender, EventArgs e)
        {
            HiddenAll();
            log.Visibility = Visibility.Visible;

        }

        private async void HandleTaskExcute(object sender, TaskModel task)
        {
            var sqlServer = new TaskRepository();
            var existingModel = await sqlServer.GetOperationByIdAsync(Convert.ToInt32(task.ID));
            Global.OperationModel = existingModel;
            if (existingModel != null)
            {
                if (existingModel.DataReady)
                {
                     
                    var filterCond = new FileterConditionModel();
                    filterCond.PN = existingModel.PN;
                    filterCond.ON = existingModel.SN;
                    filterCond.StartDateTime = existingModel.StartDateTime;
                    filterCond.StopDateTime = existingModel.EndDateTime;

                    string[] conditions = existingModel.Condition.Trim(',').Split(',');
                    string[] vd_vgs = conditions.ElementAt(0).Split(';');

                    foreach (var item in vd_vgs)
                    {
                        filterCond.VD_VG_Conditon.Add(item.Trim(';'));
                    }
                    filterCond.Min = Convert.ToDouble(conditions.ElementAt(1));
                    filterCond.Max = Convert.ToDouble(conditions.ElementAt(2));
                    int count = conditions.Count();
                    for (int i = 3; i < count; i++)
                    {
                        filterCond.FreqBands.Add(conditions.ElementAt(i));

                    }

                    //await Task.Run(() =>
                    //{
                    //    filter.SetFileterCondition(filterCond);
                    //});

                    filter.SetFileterCondition(filterCond);
                    await Task.Run(() =>
                    {
                        filter.Btn_Next_Clicked(null, null);
                    });
                    //filter.Btn_Next_Clicked(null, null);

                    //filter.Btn_Calcute_Click(null, null);
                    await Task.Run(() =>
                    {
                        filter.Btn_Calcute_Click(null, null);
                    });
                    vm.ContentTitle = Global.TaskModel.TaskName +"-" + "Amplifier" +"-" + "MML806";


                }


            }
            _task = task;
            HandleOperationBtnClick(null, null);

        }


       
        private void HandleAddBtnClick(object sender, EventArgs e)
        {
            HiddenAll();
            addPage.Visibility = Visibility.Visible;
            Console.WriteLine("This is add button click event.");
        }
        private void HiddenAll()
        {
            treeView.Visibility = Visibility.Collapsed;
            separator.Visibility = Visibility.Collapsed;
            welComeStackPanel.Visibility = Visibility.Collapsed;
            home.Visibility = Visibility.Collapsed;
            contentGrid.Visibility = Visibility.Collapsed;
            log.Visibility = Visibility.Collapsed;
            addPage.Visibility = Visibility.Collapsed;
        }

        private void Menu_Preview_Clicked(object sender, RoutedEventArgs e)
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string pptFile = System.IO.Path.Combine(appDir, "resources", "files", "T_MML806_V3.pdf");
            try
            {
                var pdfShown = new PdfShowWin();
                pdfShown.ShowPdf(pptFile);
                pdfShown.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        
    }

    internal class MainWindowModel : ObeservableObject
    {
        private NewTreeViewItem record;
        public NewTreeViewItem Record
        {
            get { return record; }
            set { this.record = value; RaisePropertyChanged(nameof(Record)); }
        }

        public ObservableCollection<NewTreeViewItem> Records { set; get; } = new ObservableCollection<NewTreeViewItem>();

        private bool popupVisible;
        public bool PopupVisible
        {
            get { return popupVisible; }

            set { this.popupVisible = value; RaisePropertyChanged(nameof(PopupVisible)); }
        }

        private string popupText;
        public string PopupText
        {
            get { return popupText; }
            set {  this.popupText = value; RaisePropertyChanged(nameof(PopupText)); }
        }

        private string popupTitle;
        public string PopupTitle
        {
            get { return popupTitle; }
            set { this.popupTitle = value; RaisePropertyChanged(nameof(PopupTitle)); }
        }

        private string contentTitle;
        public string ContentTitle
        {
            get { return contentTitle; }
            set { this.contentTitle = value; RaisePropertyChanged(nameof(ContentTitle)); }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set { this._isBusy = value; RaisePropertyChanged(nameof(IsBusy)); }
        }


        private string busyMessage;
        public string BusyMessage
        {
            get { return busyMessage; }
            set { this.busyMessage = value; RaisePropertyChanged(nameof(BusyMessage)); }
        }

        private string _logText;
        public string LogText
        {
            get { return _logText; }
            set { this._logText = value; RaisePropertyChanged(nameof(LogText)); }
        }
    }

    public class NewTreeViewItem : ObeservableObject
    {
        private string content;
        public string Content
        {
            get { return content; }
            set { this.content = value; RaisePropertyChanged(nameof(content)); }
        }
        public NewTreeViewItem Parent { get; set; }

        public bool IsLeaf
        {
            // ��� Children ����Ϊ�գ��� Count Ϊ 0������ΪҶ�ӽڵ�
            get { return Children == null || Children.Count == 0; }
        }

        private bool visible = false;
        public bool Visible
        {
            get { return visible; }
            set { this.visible = value; RaisePropertyChanged(nameof(visible)); }
        }
        public ObservableCollection<NewTreeViewItem> Children { get; set; } = new ObservableCollection<NewTreeViewItem>();

        public NewTreeViewItem()
        {
            Children.CollectionChanged += (sender, e) => {
                // ���Ӽ��ϱ仯ʱ��֪ͨ��ϵͳ IsLeaf ���Կ����Ѹı�
                RaisePropertyChanged(nameof(IsLeaf));
            };
        }
    }


    public static class TreeExtensions
    {
        public static void AddChild(this NewTreeViewItem parent, NewTreeViewItem child)
        {
            child.Parent = parent;
            parent.Children.Add(child);
        }

        public static void AddChildren(this NewTreeViewItem parent, IEnumerable<string> contents)
        {
            foreach (var content in contents)
            {
                parent.AddChild(new NewTreeViewItem { Content = content });
            }
        }
    }

}

