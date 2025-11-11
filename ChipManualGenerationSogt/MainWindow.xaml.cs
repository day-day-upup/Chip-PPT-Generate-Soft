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
using System.Text.Json;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using CommunityToolkit.Mvvm.ComponentModel;
namespace ChipManualGenerationSogt
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        MainWindowModel vm;
        PptDataModel pptDataModel;
       
        List<PlotModel> plots;
        User _user;
        FileterConditionModel _filterCondition;
        List<string> _plotNames = new List<string>();
        string _fileFolerPath = "";
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
            //vm.LogText = "这是一个日志测试:A Log in\n 这是一个日志测试:A Select Amplifier MM809\n  这是一个日志测试:A Enter SN:L004x,ON:L004x\n 这是一个日志测试： A Select contion:VG 4.0V,VD 1.5V\n 这是一个日志测试： A Select Filter:MML806\n 这是一个日志测试： A Select Amplifier MM809\n 这是一个日志测试： A Select contion:VG 4.0V,VD 1.5V\n 这是一个日志测试： A Select Filter:MML806\n 这是一个日志测试： A Select Amplifier MM809\n 这是一个日志测试： A Select contion:VG 4.0V,VD 1.5V\n 这是一个日志测试： A Select Filter:MML806\n 这是一个日志测试： A Select Amplifier MM809\n 这是一个日志测试： A Select contion:VG 4.0V,Idd:67mA\n 这是一个日志测试： A Select Filter:M";
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
            taskMangeControls.AddEvent += (sender, e) =>
            {
                HiddenAll();
                newTaskPage.Visibility = Visibility.Visible;
            };
            taskMangeControls.QueryLogEvent += (sender, e) =>
            {
                HandleLogBtnClick(null, null);
            };
            taskMangeControls.DetailShowEvent += (sender, e) =>
            {
                var taskItem = e as TaskTableItem;
                if (taskItem != null)
                { 

                
                }
                HiddenAll();
                newTaskPage.Visibility = Visibility.Visible;
                newTaskWin.ShowCurrentTaskConfigure(taskItem);

            };
            newTaskWin.BackEvent += (sender, e) =>
            {
                //newTaskPage.Visibility = Visibility.Hidden;
                HiddenAll();
                taskMangeControls.RefreshTask();
                home.Visibility = Visibility.Visible;
            };

        
         

            logWin.BackEvent += (sender, e) =>
            {

                HiddenAll();
                //taskMangeControls.RefreshTask();
                home.Visibility = Visibility.Visible;
            };

            var root1 = new DeviceTreeViewItemModel("Amplifier");

            // 创建 Root 1 的子节点
            //root1.Children.Add(new DeviceTreeViewItemModel("Low Noise Amplifier"));
            //root1.Children.Add(new DeviceTreeViewItemModel("Power Amplifier"));

            var subRoot1 = new DeviceTreeViewItemModel("Low Noise Amplifier");
            var item = new DeviceTreeViewItemModel("MML");
            item.IsSelectedInModel = true;
            subRoot1.Children.Add(item);
            //subRoot1.Children.Add(new DeviceTreeViewItemModel("A-Sub-X2"));

            // 将子树添加到 Root 1
            root1.Children.Add(subRoot1);
          

            // 创建 Root 2 的子节点
            //root2.Children.Add(new DeviceTreeViewItemModel("Actuator Model B1"));
            //root2.Children.Add(new DeviceTreeViewItemModel("Actuator Model B2"));
            vm.TreeViewSources.Add(root1);

            //var win = new LoginW();
            //win.Show();

            //vm.Records.Add(new NewTreeViewItem
            //{
            //    Content = "Amplifier",
            //    Visible = true,
            //    // Children.Add(new NewTreeViewItem { Content = "人民日报" }),
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
            //pdfShown.ShowPdf("C:\\Users\\pengyang\\Desktop\\books\\UNIX编程艺术.pdf");
            //pdfShown.Show();

            //test3();
            //PPTChange();
        }
        public void PptDataModeFactory()
        {
            var imgs = images.GetAllImage();
            //********************母版信息
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
            #region 第一页
            string features = "Features\n";
            foreach (var item in table.GetFeatureTableData())
            {
                features += "\u2022" +"    " +item.name + " : " + item.info + "\n";
            }
            features = features.TrimEnd('\n');

            //这个的匹配赋值方式可能还需要更改

            string elecCondition = "TA = +25\u2103, " + _filterCondition.VD_VG_Conditon.ElementAt(0).Replace('&', ',') + " Typical";

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



            #region 曲线页
            // 这种划分方式参考了MML806_V3.pptx
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


            #region 倒数第五页

            pptDataModel.EndToFront5Page = new EndToFront5();
            pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTableTitle = "Absolute Maximum Ratings";
            pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTable = DataConverter.ConvertListToTwoDArray(table.GetAbsoluteRatingsData());
            pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTableTitle = "Typical Supply Current";

            string BasePath = "F:\\PROJECT\\ChipManualGeneration\\exe\\app\\ChipManualGenerationSogt\\bin\\Debug\\resources\\files";

            pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTable = DataConverter.ConvertThreeElementListToTwoDArray(table.GetCurrentVdVgData());
            pptDataModel.EndToFront5Page.WarningImage = new ImageModel
            {
                ImagePath = System.IO.Path.Combine(BasePath, "防静电标志.jpg"),
                Height = 50_0000,
                Width = 50_0000,
            };
            pptDataModel.EndToFront5Page.WarningText = "ELECTROSTATIC SENSITIVE DEVICE\n OBSERVE HANDLING PRECAUTIONS";
            #endregion

            #region  倒数第4页

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

            #region 倒数第3页
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

            #region 倒数第2页
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


            #region 倒数第1页
            pptDataModel.LastPage = new LastPage();

            pptDataModel.LastPage.Image = new ImageModel
            {
                ImageName = imgs.ElementAt(4).name,
                ImagePath = imgs.ElementAt(4).filePath,
                Height = 500_0000,
                Width = 150_0000
            };
            pptDataModel.LastPage.Text1 = "Direct Mounting\n1.Typically, the die is mounted directly on the ground plane.\n2.If the thickness difference between the substrate (thickness c) and the die (thickness d) exceeds 0.05 mm (i.e., c – d > 0.05 mm), it is recommended to first mount the die on a heat spreader, then attach the heat spreader to the ground plane.\n3.Heat Spreader Material: Molybdenum-copper (MoCu) alloy is commonly used.\n4.Heat Sink Thickness (b): Should be within the range of (c – d – 0.05 mm) to (c –d + 0.05 mm).\n5.Spacing (a): The gap between the bare die and the 50Ω transmission line should typically be 0.05 mm to 0.1 mm. If the application frequency is higher than 40GHz, then this gap is recommended to be 0.05mm\n Wire Bonding Interconnection\nThe connection between the die and the 50Ω transmission line is usually made using 25 \u03BCm diameter gold (Au) wires, bonded via wedge bonding or ball bonding processes.\nDie Attachment Methods\n1.Conductive Epoxy:\nAfter adhesive application, cure according to the manufacturer’s recommended temperature profile.\n2.Au-Sn80/20 Eutectic Bonding:\nUse preformed Au-Sn80/20 solder preforms.\nPerform bonding in an inert atmosphere (N\u2082 or forming gas: 90% N\u2082 + 10% H\u2082).\nKeep the time above 320\u2103 to less than 20 seconds to prevent excessive intermetallic formation.\n";
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

            // 可选：自定义日志输出（例如写入文件）
            // copier.Log = msg => File.AppendAllText("copy.log", $"{DateTime.Now:HH:mm:ss} {msg}\n");

            try
            {
                //更具ON查二级路径
                copier.CopyMatchingSubFolders(
                    networkRoot: @"\\DATAPC03\RFAutoTestReport$\Chip Verification",
                    PN: keyword,
                    ON: sn,
                    localTargetBase: System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiedReports"),
                    username: "",   // 留空表示使用当前用户凭据
                    password: ""
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"?? 程序异常: {ex.Message}");
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

            // 1. 定义所有要处理的属性及其对应的目标列表
            // 使用元组数组来映射来源字典和目标列表
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
            // 2. 遍历所有映射并统一处理
            foreach (var mapping in fileMappings)
            {
                // 尝试从当前来源字典中获取文件列表
                if (mapping.sourceDict.TryGetValue(tempKey, out var files))
                {
                    // 使用通用的筛选函数获取匹配的文件
                    var matchingFiles = GetMatchingFiles(files, conditions);

                    // 将结果一次性添加到目标列表中
                    mapping.targetList.AddRange(matchingFiles);
                }
            }


            GeneratePlotsByTemperaturen(filesModel25, condition.Min, condition.Max, LegendTextType.Elec);
            var tem1 = "Measurement Plots: S-parameters\n" + "TA = +25\u2103";
            var tem2 = "Measurement Plots: P1dB\n" + "TA = +25\u2103";
            var tem3 = "Measurement Plots: OIP3\n" + "TA = +25\u2103";
            var tem4 = "Measurement Plots: PSAT\n" + "TA = +25\u2103";
            var tem5 = "Measurement Plots: Noise Figure\n" + "TA = +25\u2103";
            _plotNames.Add(tem1);
            _plotNames.Add(tem2);
            _plotNames.Add(tem3);
            _plotNames.Add(tem4);
            _plotNames.Add(tem5);

            List<Temperature25FilePathModel> filesModelVgs = new List<Temperature25FilePathModel>();

     

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
            foreach (var item in filesModelVgs)// 不同的VD下， 不同的三温度图
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


               for(int i=0; i< filesModelVgs.Count; i++)// 不同的VD下， 不同的三温度图
            {
                if (filesModelVgs.ElementAt(i).SList.Count > 0)
                {
                    GeneratePlotsByTemperaturen(filesModelVgs.ElementAt(i), condition.Min, condition.Max, LegendTextType.Temp);
                    // 这个数组用于处理参数列表， 是用来计算的

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
        // 通用辅助函数：处理所有需要分割匹配条件的逻辑
        // ----------------------------------------------------------------------
        private IEnumerable<string> GetMatchingFiles(
            IEnumerable<string> files,
            IEnumerable<string> conditions)
        {
            // 使用 LINQ 查找所有匹配的文件
            return files.Where(item =>
                conditions.Any(subCon =>
                {
                    // 确定分隔符：如果包含 '&' 则用 '&'，否则用 ','
                    char splitChar = subCon.Contains('&') ? '&' : ',';

                    // 分割字符串，并获取第一个部分作为匹配条件
                    // 使用 FirstOrDefault() 替代 Split().FirstOrDefault()，性能更优
                    string matchPart = subCon.Split(splitChar).FirstOrDefault();

                    // 检查 matchPart 不为空且 item（文件路径）包含匹配部分
                    return !string.IsNullOrEmpty(matchPart) && item.Contains(matchPart);
                }));
        }
        public List<string> SelectUnique5VFilePerTemperature(List<string> allFiles, string vd)
        {
            var selectedFiles = allFiles
                // 1. Where (筛选)：只保留 VD=5V 的文件
                .Where(filePath => ExtractVoltageKey(filePath) == vd)

                // 2. GroupBy (分组)：按温度（例如 "25.0deg"）进行分组
                .GroupBy(filePath => ExtractTemperatureKey(filePath))

                // 3. Select (选择)：从每个温度组中选择第一个文件路径
                .Select(group => group.First())

                // 4. ToList：转换为最终列表
                .ToList();

            return selectedFiles;
        }

        // 1. 提取温度键 (例如 "25.0deg", "-40.0deg")
        private string ExtractTemperatureKey(string filePath)
        {
            // 假设温度总是在文件名末尾的 "_XX.Xdeg_" 格式之前
            // 使用正则表达式查找数字、小数点、可选的负号和 "deg"
            var match = Regex.Match(filePath, @"_(-?\d+\.\d+deg)_");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return "UNKNOWN_TEMP";
        }

        // 2. 提取电压键 (例如 "VD=5V")
        private string ExtractVoltageKey(string filePath)
        {


            // 找到路径中的 VD/ID 段，然后提取 VD 的值。
            string normalizedPath = filePath.Replace('\\', '/');
            string[] pathSegments = normalizedPath.Split('/');

            string vdIdSegment = pathSegments.FirstOrDefault(s => s.StartsWith("-VD=") || s.StartsWith("VD="));

            if (vdIdSegment != null)
            {
                // 使用正则提取 VD 的值，例如从 "-VD=5V&ID=90mA" 中提取 "5V"
                var match = Regex.Match(vdIdSegment, @"-?VD=(\d+V)");
                //if (match.Success)
                //{
                //    return match.Groups[1].Value; // 提取 "5V", "4V" 等
                //}
                //if (match.Success)
                //{
                //    return match.Groups[1].Value; // 提取 "5V", "4V" 等
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

            // 提取温度（以 "deg" 结尾的段）
            temperature = parts.FirstOrDefault(p => p.EndsWith("deg", StringComparison.OrdinalIgnoreCase))
                              ?? "UnknownTemp";

            // 提取电气参数（包含 "VD=" 的段）
            elecParam = parts.FirstOrDefault(p => p.Contains("VD="))
                            ?? "UnknownParam";
        }
        //private void btn_Generate_Click(object sender, RoutedEventArgs e)
        private void GeneratePlots(AmpfilierFilesbyGroup ampParams, FileterConditionModel condition)
        {

            string filePath = "";

            //#region 读取标准S参数

            // 获得标准S参数
            if (ampParams.DataSparaFilePaths.Count > 4)
            {
                filePath = ampParams.DataSparaFilePaths[3];
            }
            else
            {

                filePath = ampParams.DataSparaFilePaths[0];
            }


            if (!ampParams.DataSparabyTemp.TryGetValue("25.0deg", out var files))
            {
                // 如果 files 为空或未找到键，则返回或抛出异常
                return; // 或者 throw new KeyNotFoundException("未找到 '25.0deg' 的数据。");
            }

            // 1. 使用 LINQ 查找第一个匹配的文件路径
            string firstMatchingFile = files.FirstOrDefault(item =>
            {
                // 遍历所有条件，找到第一个符合的文件即返回
                return condition.VD_VG_Conditon.Any(subCon =>
                {
                    // 优化点：使用 char.Split 代替 string.Split
                    char splitChar = subCon.Contains('&') ? '&' : ',';

                    // 分割字符串，并获取第一个部分作为匹配条件
                    // 优化：使用 subCon.Split(splitChar).FirstOrDefault() 代替创建整个数组再访问 [0]
                    string matchPart = subCon.Split(splitChar).FirstOrDefault();

                    // 检查 item（文件路径）是否包含匹配部分
                    return matchPart != null && item.Contains(matchPart);
                });
            });


            // 2. 检查是否找到文件，并进行绘图
            if (firstMatchingFile != null)
            {
                // 将单个文件放入 List<string> (files2)
                List<string> files2 = new List<string> { firstMatchingFile };

                // 执行绘图操作
                GenerateSParaPlots(files2, condition.Min, condition.Max, LegendTextType.Elec);
            }

        }


        private void GeneratePlots(AmpfilierFilesbyGroup ampParams, FileterConditionModel condition,string key)
        {

            string filePath = "";

            //#region 读取标准S参数

            // 获得标准S参数
            if (ampParams.DataSparaFilePaths.Count > 4)
            {
                filePath = ampParams.DataSparaFilePaths[3];
            }
            else
            {

                filePath = ampParams.DataSparaFilePaths[0];
            }


            if (!ampParams.DataSparabyTemp.TryGetValue("25.0deg", out var files))
            {
                // 如果 files 为空或未找到键，则返回或抛出异常
                return; // 或者 throw new KeyNotFoundException("未找到 '25.0deg' 的数据。");
            }

          

           // string finalMatchPart = null;

            //if (!string.IsNullOrEmpty(key))
            //{
            //    // 2. 应用原有的复杂规则来提取匹配部分 (matchPart)

            //    // 确定分隔符：检查是否包含 '&'，否则使用 ','
            //    char splitChar = key.Contains('&') ? '&' : ',';

            //    // 分割字符串，并获取第一个部分作为最终的匹配条件
            //    finalMatchPart = key.Split(splitChar).FirstOrDefault();
            //}

            // 1. 使用 LINQ 查找第一个匹配的文件路径
            string firstMatchingFile = files.FirstOrDefault(item =>
            {
                string secondMatchPart = null;
                if (key.Contains('&'))
                {
                    secondMatchPart = key.Replace('&', ',');
                }
                else
                {
                    secondMatchPart = key.Replace(',', '&');
                }

                    return !string.IsNullOrEmpty(key) && (item.Contains(key) || item.Contains(secondMatchPart));
            });


            // 2. 检查是否找到文件，并进行绘图
            if (firstMatchingFile != null)
            {
                // 将单个文件放入 List<string> (files2)
                List<string> files2 = new List<string> { firstMatchingFile };

                // 执行绘图操作
                GenerateSParaPlots(files2, condition.Min, condition.Max, LegendTextType.Elec);
            }

        }


        private void GenerateSParaPlots(List<string> files, double xMin, double xMax, LegendTextType type)
        {

            var plotS21 = new PlotModel();
            plotS21.YLabel = "GAIN(dB)";
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
                    case LegendTextType.Temp:// legend 用温度当不同
                        curveS11.Legend = temperature;
                        curveS12.Legend = temperature;
                        curveS21.Legend = temperature;
                        curveS22.Legend = temperature;
                        break;

                    case LegendTextType.Elec:// 以电气参数当不同
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


            //重新调整。 让x轴包边，最小值在做左边刻度线处， 最大值在最右边刻度线出。 采用的方法是调整每隔增量， 选择10附近能整除的
            if (PlotModel.CalculateFixedInterval((int)xMin, (int)xMax, 10, out int xInterval))
            {
                plotS11.xAxisInterval = xInterval;
                plotS12.xAxisInterval = xInterval;
                plotS21.xAxisInterval = xInterval;
                plotS22.xAxisInterval = xInterval;
            }
            double newYMin, newYMax;
            double yInterval;
            const int TargetDivisions = 10; // 目标 10 个刻度

            //PlotModel.CalculateNiceRange((double)plotS11.yMin, (double)plotS11.yMax, TargetDivisions,
            //                                      out newYMin, out newYMax, out yInterval);
            PlotModel.S11CalculateNiceRange((double)plotS11.yMin, 0, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS11.yMin = newYMin;
            plotS11.yMax = newYMax;
            plotS11.yAxisInterval = yInterval;
            plotS11.Alignment = ScottPlot.Alignment.LowerRight;
            //plotS11.Alignment = ScottPlot.Alignment.UpperRight;


            //PlotModel.CalculateNiceRange((double)plotS12.yMin, (double)plotS12.yMax, TargetDivisions,
            //                                      out newYMin, out newYMax, out yInterval);


            PlotModel.S11CalculateNiceRange((double)plotS12.yMin, (double)plotS12.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS12.yMin = newYMin;
            plotS12.yMax = newYMax;
            plotS12.yAxisInterval = yInterval;
            //plotS12.Alignment = ScottPlot.Alignment.UpperRight;


            //PlotModel.CalculateNiceRange((double)plotS21.yMin, (double)plotS21.yMax, TargetDivisions,
            //                                      out newYMin, out newYMax, out yInterval);

            PlotModel.S21CalculateNiceRange((double)plotS21.yMin, (double)plotS21.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS21.yMin = newYMin;
            plotS21.yMax = newYMax;
            plotS21.yAxisInterval = yInterval;
            plotS21.Alignment = ScottPlot.Alignment.LowerRight;
            //plotS11.Alignment = ScottPlot.Alignment.LowerRight;

            //PlotModel.CalculateNiceRange((double)plotS22.yMin, (double)plotS22.yMax, TargetDivisions,
            //                                     out newYMin, out newYMax, out yInterval);
            PlotModel.S11CalculateNiceRange((double)plotS22.yMin, (double)plotS22.yMax, TargetDivisions,
                                                 out newYMin, out newYMax, out yInterval);
            plotS22.yMin = newYMin;
            plotS22.yMax = newYMax;
            plotS22.yAxisInterval = yInterval;
            plotS22.Alignment = ScottPlot.Alignment.LowerRight;


            curves.AddPlot(plotS21);
            curves.AddPlot(plotS11);
            curves.AddPlot(plotS12);
            curves.AddPlot(plotS22);
        }

        /// <summary>
        /// 用来产生NF, Pxdb, OIP3, Psat这几个图像
        /// </summary>
        /// <param name="files"></param>
        /// <param name="xMin"></param>
        /// <param name="xMax"></param>
        /// <param name="type"></param>
        private void GeneratePlotsByTemperaturen(Temperature25FilePathModel filesModel25, double xMin, double xMax, LegendTextType type)

        {

            //PlotModel.Init();
            #region S参数处理
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
            plotS21.YLabel = "GAIN(dB)";
            plotS21.XLabel = "FREQUENCY(GHz)";

            ////////////s22
            var plotS22 = new PlotModel();
            plotS22.YLabel = "OUTPUT RETURN LOSS(dB)";
            plotS22.XLabel = "FREQUENCY(GHz)";
            var analyzer = new S2PParser();

            string oldChar = "deg";
            string newChar = "\u00B0C"; 

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
                    case LegendTextType.Temp:// legend 用温度当不同
                        curveS11.Legend = temperature.Replace(oldChar, newChar);
                        curveS12.Legend = temperature.Replace(oldChar, newChar);
                        curveS21.Legend = temperature.Replace(oldChar, newChar);
                        curveS22.Legend = temperature.Replace(oldChar, newChar);
                        break;

                    case LegendTextType.Elec:// 以电气参数当不同
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


            //重新调整。 让x轴包边，最小值在做左边刻度线处， 最大值在最右边刻度线出。 采用的方法是调整每隔增量， 选择10附近能整除的
            if (PlotModel.CalculateFixedInterval((int)xMin, (int)xMax, 10, out int xInterval))
            {
                plotS11.xAxisInterval = xInterval;
                plotS12.xAxisInterval = xInterval;
                plotS21.xAxisInterval = xInterval;
                plotS22.xAxisInterval = xInterval;
            }
            double newYMin, newYMax;
            double yInterval;
            const int TargetDivisions = 10; // 目标 10 个刻度

            PlotModel.S11CalculateNiceRange((double)plotS11.yMin, (double)plotS11.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS11.yMin = newYMin;
            plotS11.yMax = newYMax;
            plotS11.yAxisInterval = yInterval;
            plotS11.Alignment = ScottPlot.Alignment.LowerRight;


            PlotModel.S11CalculateNiceRange((double)plotS12.yMin, (double)plotS12.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS12.yMin = newYMin;
            plotS12.yMax = newYMax;
            plotS12.yAxisInterval = yInterval;
            //plotS12.Alignment = ScottPlot.Alignment.UpperRight;


            PlotModel.S21CalculateNiceRange((double)plotS21.yMin, (double)plotS21.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            plotS21.yMin = newYMin;
            plotS21.yMax = newYMax;
            plotS21.yAxisInterval = yInterval;
            plotS21.Alignment = ScottPlot.Alignment.LowerRight;


            PlotModel.S11CalculateNiceRange((double)plotS22.yMin, (double)plotS22.yMax, TargetDivisions,
                                                 out newYMin, out newYMax, out yInterval);
            plotS22.yMin = newYMin;
            plotS22.yMax = newYMax;
            plotS22.yAxisInterval = yInterval;
            plotS22.Alignment = ScottPlot.Alignment.LowerRight;


            curves.AddPlot(plotS21);
            curves.AddPlot(plotS11);
            curves.AddPlot(plotS12);
            curves.AddPlot(plotS22);

            #endregion


            #region 其他参数处理
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
                        curve.Legend = temperature.Replace(oldChar, newChar);
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
            

            //PlotModel.CalculateNiceRange((double)plotP1db.yMin, (double)plotP1db.yMax, TargetDivisions,
            //                                      out newYMin, out newYMax, out yInterval);
            PlotModel.PxdbCalculateNiceRange((double)plotP1db.yMin, (double)plotP1db.yMax, TargetDivisions,
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
                        curve.Legend = temperature.Replace(oldChar, newChar);
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


            PlotModel.S21CalculateNiceRange((double)plotOIP3.yMin, (double)plotOIP3.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            
            plotOIP3.yMin = newYMin;
            plotOIP3.yMax = newYMax;
            plotOIP3.yAxisInterval = yInterval;


            ////////////psat
            var plotPsat = new PlotModel();
            plotPsat.YLabel = "PSAT(dBm)";
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
                        curve.Legend = temperature.Replace(oldChar, newChar);
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


            PlotModel.PsatCalculateNiceRange((double)plotPsat.yMin, (double)plotPsat.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);
            
            plotPsat.yMin = newYMin;
            plotPsat.yMax = newYMax;
            plotPsat.yAxisInterval = yInterval;

            var plotNF = new PlotModel();
            plotNF.YLabel = "NOISE FIGURE(dB)";
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
                        curve.Legend = temperature.Replace(oldChar, newChar);
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


            PlotModel.NFCalculateNiceRange((double)plotNF.yMin, (double)plotNF.yMax, TargetDivisions,
                                                  out newYMin, out newYMax, out yInterval);

            plotNF.yMin = newYMin;
            plotNF.yMax = newYMax;
            plotNF.yAxisInterval = yInterval;
            plotNF.Alignment = ScottPlot.Alignment.UpperRight;




            curves.AddPlot(plotP1db);
            curves.AddPlot(plotOIP3);
            curves.AddPlot(plotPsat);
            curves.AddPlot(plotNF);
            #endregion
        }

        

        private void CalcuteParameter(List<string> files, List<string> freqBands)
        {
            // files 0-s参数文件，1-pxdb，2-oip3，3-psat，4-nf
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
            foreach (var band in freqBands)// 遍历所有频段， 计算频段的min， type， amx
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
        /// <param name="band">包含两个数字的字符串，中间可能有任意分隔符。</param>
        /// <param name="minValue">输出：解析出的最小值。</param>
        /// <param name="maxValue">输出：解析出的最大值。</param>
        /// <returns>如果成功解析出两个数字，则返回 true；否则返回 false。</returns>
        public bool TryParseMinMaxBand(string band, out double minValue, out double maxValue)
        {
            minValue = 0.0;
            maxValue = 0.0;

            if (string.IsNullOrWhiteSpace(band))
            {
                return false;
            }

            // 优化后的正则表达式：只匹配非负数（数字可选带小数点）
            // 关键改变：移除了前面的 "-?"
            // \d+ : 匹配一个或多个数字 (0-9)
            // \.? : 匹配可选的小数点
            // \d* : 匹配 0 个或多个数字（用于支持 .5, 10.0 等）
            const string pattern = @"\d+\.?\d*";

            // 1. 查找所有匹配非负数字的部分
            // 注意：这里的破折号、逗号等现在只被视为分隔符
            MatchCollection matches = Regex.Matches(band, pattern);

            // 2. 检查是否找到了至少两个数字
            if (matches.Count >= 2)
            {
                // 3. 提取前两个匹配到的数字
                if (double.TryParse(matches[0].Value, out double num1) &&
                    double.TryParse(matches[1].Value, out double num2))
                {
                    // 4. 确定最小值和最大值 (因为输入顺序不一定确定)
                    minValue = Math.Min(num1, num2);
                    maxValue = Math.Max(num1, num2);
                    return true;
                }
            }

            // 如果没有找到两个数字，或者转换失败
            return false;
        }
        public XYPoint FilterXYPointData(XYPoint sourcePoint, double min, double max)
        {
            // 1. 基础检查：确保源对象、XArrys 和 YArrys 都存在且长度匹配
            if (sourcePoint == null ||
                sourcePoint.XArrys == null ||
                sourcePoint.YArrys == null ||
                sourcePoint.XArrys.Length != sourcePoint.YArrys.Length)
            {
                return null; // 数据无效，返回 null 或抛出异常
            }

            // 2. 使用 Zip 方法将 X 数组和 Y 数组按索引配对成匿名对象 (X, Y)
            var pairedData = sourcePoint.XArrys.Zip(sourcePoint.YArrys, (x, y) => new { X = x, Y = y });

            // 3. 筛选：只保留 X 值在 [min, max] 范围内的对
            var filteredPairs = pairedData
                .Where(pair => pair.X >= min && pair.X <= max)
                .ToList();

            // 4. 创建新的 XYPoint 对象，并拆分回 X 数组和 Y 数组
            return new XYPoint
            {
                // 新的 Size 可以基于筛选后的数量，或者保持原有的 Size
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
            // 3. 将子节点集合赋给父节点
            parentItem.Children = children;

            // 4. 添加到根集合
            vm.Records.Add(parentItem);
            //vm.Records.Add(new NewTreeViewItem
            //{
            //    Content = "Amplifier",
            //    Visible = true,
            //    // Children.Add(new NewTreeViewItem { Content = "人民日报" }),
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
                Console.WriteLine(item.Content); // ? 安全访问
                PopupTitleChange(item.Content);
                //item.Children.Add(new NewTreeViewItem { Content = "MML808" });
                vm.Record = item;
                vm.PopupVisible = true;
                // 你也可以赋值给 vm.Record（如果需要）
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
            var selectedItem = e.NewValue as NewTreeViewItem; // ?? 你绑定的数据类型

            if (selectedItem != null)
            {
                try
                {
                    if (selectedItem.Parent != null)
                    {
                        Console.WriteLine($"你点击了：{selectedItem.Content} {selectedItem.Parent.Content}");
                        // 你可以在这里更新 ViewModel、弹出菜单、显示详情等
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
            //    //保存 Amplifier下的序列到配置文件
            //    if (item.Content == "Amplifier")
            //    {
            //        Properties.Settings.Default.Amplifier.Clear();
            //        foreach (var child in item.Children)
            //        {
            //            Properties.Settings.Default.Amplifier.Add(child.Content);
            //        }
            //    }
            //}
            //Properties.Settings.Default.Save(); // 必须调用 Save() 才会写入磁盘！
            base.OnClosed(e);
        }

        private void PopupTitleChange(string title)
        {
            vm.PopupTitle = $"Please Add A Chip Serial Number of {title}";

        }


        private bool PPTChange(string filePath = @"resources\files\T_MML806_V3.pptx", string tagetFilePath = @"resources\files\demo.pptx")
        {
            //filePath = @"F:\PROJECT\ChipManualGeneration\exe\T_MML806_V3.pptx";
            tagetFilePath = @"F:\PROJECT\ChipManualGeneration\exe\ChipManualGenerationSogt\ChipManualGenerationSogt\bin\Debug\resources\files\demo.pptx";
            bool success = false;
            try
            {

                if (!File.Exists(filePath))
                    throw new FileNotFoundException("未找到PPT文件", filePath);
                if (File.Exists(tagetFilePath))
                {
                    try
                    {
                        File.Delete(tagetFilePath);
                    }
                    catch (IOException ex)
                    {
                        throw new IOException($"Canno't delete file：{tagetFilePath}", ex);
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
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{right bar info}", "GaAs Low Noise Amplifier MMIC 45 – 90GHz");
                    var presentationPart = presentationDoc.PresentationPart;

                    // 获取第一个 slide part（通过关系）
                    var slideId = presentationPart.Presentation.SlideIdList?.GetFirstChild<P.SlideId>();
                    //var slideId = presentationPart.Presentation.SlideIdList?.FirstOrDefault();
                    if (slideId == null)
                        throw new InvalidOperationException("PPT中没有幻灯片。");


                    // ? 关键：v3.3.0 中用 GetPartById 获取 SlidePart（类型是 OpenXmlPart，但可转为 Slide）
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    var slide = slidePart.Slide;

                    var currentY = 1214000L; // 初始 Y 位置
                    const long verticalSpacing = 500000; // 间距 100,000 EMU

                    string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ±2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50Ω\nChip Size: 1.766 x 2.0 x 0.05mm";
                    string x = "Feautures\n" + table.GetFeatureTableInfo();
                    // 添加文本框并获取实际高度
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
                            { "Gain Flatness",  "", "±1.0", "", "", "±1.0", "", "dB" },
                            { "Noise Figure",  "", "±1.0", "", "", "±1.0", "", "dB" },
                            { "P1dB - Output 1dB Compression",  "", "12", "", "", "14", "", "dBm" },
                            { "Psat - Saturated Output Power",  "", "12", "", "", "14", "", "dBm" },
                            { "OIP3 - Output Third Order Intercept",  "", "12", "", "", "14", "", "dBm" },
                            { "Input Return Loss",  "", "12", "", "", "14", "", "dB" },
                            { "Output Return Loss",  "", "12", "", "", "14", "", "dB" }
                    };

                    PptModifier.AddTable(slide, table.GetParameterTableInfo(), 914400, currentY, 6000000, 3800000);
                    currentY += 2000000 + verticalSpacing; // 表格高度 + 间距


                    // 添加新幻灯片
                    var newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    //AddNewSlideFromLayout

                    var newSlide = newSlidePart.Slide;
                    // 示例：在新幻灯片上添加表格
                    int originX = 914400;

                    int originY = 1314000;
                    currentY = originY;
                    string info = "Measurement Plots: S-parameters\n TA = +25\u2103";// \u2103 是摄氏度的符号
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
                    info = "Absolute Maximum Ratings";// \u2103 是摄氏度的符号
                    height = PptModifier.AddTextBoxUnderline(newSlide, info, originX, originY, 3500_000, 300000);

                    info = "Typical Supply Current vs. VD,VG";// \u2103 是摄氏度的符号
                    height = PptModifier.AddTextBoxUnderline(newSlide, info, 914400 + 2700000 + 700000, originY, 3500_000, 300000);
                    currentY += height + 50000;
                    tableData = new string[,]
                            {

                            { "Drain Bias Voltage (VD)",  "+4.5V" },
                            { "Gate Bias Voltage (VG)",  "-2V to 0V" },
                            { "RF Input Power (RFIN)@(+4V)",  "+15"},
                            { "Channel Temperature",   "±1.0" },
                            { "Continuous Pdiss (T = 85 °C)\n(derate 6.1mW/°C above 85 °C)",  "12", },
                            { "Thermal Resistance\n (channel to die bottom)",   "12" },
                            { "Operating Temperature",  "55°C to +85 °C"},
                            { "Storage Temperature",  "65°C to +150 °C" },

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
                    info = "Outline Drawing: \nAll Dimensions in μm";
                    height = PptModifier.AddTextBoxCenter2(newSlide, info, originX, currentY);

                    pic1 = @"F:\PROJECT\ChipManualGeneration\exe\3.png";
                    PptModifier.AddImage(newSlidePart, pic1, 1014400, currentY + 800_000, 500_0000, 500_0000);

                    info = "Notes:\n1. Die thickness: 50μm\n2. VD bond pad is 75*75μm\u00B2 \n3. VG bond pad is 75*75μm\u00B2 \n4. RF IN/OUT bond pad is 50*86μm\u00B2 \n5. Bond pad metalization: Gold\n6. Backside metalization: Gold\n";
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
                    info = "Direct Mounting\n1.Typically, the die is mounted directly on the ground plane.\n2.If the thickness difference between the?substrate (thickness?c)?and the?die (thickness?d)?exceeds?0.05 mm (i.e.,?c?–?d?> 0.05 mm), it is recommended to first mount the die on a?heat spreader, then attach the heat spreader to the ground plane.\r\n3.Heat Spreader Material: Molybdenum-copper (MoCu) alloy is commonly used.\r\n4.Heat Sink Thickness (b): Should be within the range of?(c?–?d?– 0.05 mm)?to?(c?–?d?+ 0.05 mm).\r\n5.Spacing (a): The gap between the bare die and the 50Ω transmission line should typically be?0.05 mm to 0.1 mm.\r\nIf the application frequency is higher than 40GHz, then this gap is recommended to be 0.05mm\r\nWire Bonding Interconnection\r\nThe connection between the die and the 50Ω transmission line is usually made using?25 ?m diameter gold (Au) wires, bonded via?wedge bonding?or?ball bonding?processes.\r\nDie Attachment Methods\r\n1.Conductive Epoxy:\r\nAfter adhesive application, cure according to the manufacturer’s recommended temperature profile.\r\n2.Au-Sn80/20 Eutectic Bonding:\r\nUse preformed?Au-Sn80/20 solder preforms.\r\nPerform bonding in an inert atmosphere (N? or forming gas: 90% N? + 10% H?).\r\nKeep the time above?320°C?to?less than 20 seconds?to prevent excessive intermetallic formation.\r\n";
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
            //string pptFile = System.IO.Path.Combine(Global.FileBasePath, "demo.pptx");
            //tagetFilePath = @"F:\PROJECT\ChipManualGeneration\exe\ChipManualGenerationSogt\ChipManualGenerationSogt\bin\Debug\resources\files\demo.pptx";
            string filePath = @"resources\files\T_MML806_V3.pptx";
            bool success = false;
            try
            {

                if (!File.Exists(filePath))
                    throw new FileNotFoundException("未找到PPT文件", filePath);
                if (File.Exists(tagetFilePath))
                {
                    try
                    {
                        File.Delete(tagetFilePath);
                    }
                    catch (IOException ex)
                    {
                        throw new IOException($"Canno't delete file：{tagetFilePath}", ex);
                    }
                }
                File.Copy(filePath, tagetFilePath, overwrite: true);
                using (var presentationDoc = PresentationDocument.Open(tagetFilePath, isEditable: true))
                {
                    //修改母版的信息
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Top PN}", pptDataModel.SliderMaster.TopPN);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Version}", pptDataModel.SliderMaster.Version);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Product Name}", pptDataModel.SliderMaster.ProductName);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{Frequency Range}", pptDataModel.SliderMaster.FrequencyRange);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{R PN}", pptDataModel.SliderMaster.TopPN);
                    PptModifier.ReplaceTextSlideMasterInPresentation(presentationDoc, "{right bar info}", pptDataModel.SliderMaster.RightBarInfo);

                    var presentationPart = presentationDoc.PresentationPart;

                    // 获取第一个 slide part（通过关系）
                    var slideId = presentationPart.Presentation.SlideIdList?.GetFirstChild<P.SlideId>();
                    //var slideId = presentationPart.Presentation.SlideIdList?.FirstOrDefault();
                    if (slideId == null)
                        throw new InvalidOperationException("PPT中没有幻灯片。");

                    #region 第一页
                    // ? 关键：v3.3.0 中用 GetPartById 获取 SlidePart（类型是 OpenXmlPart，但可转为 Slide）
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    var slide = slidePart.Slide;

                    var currentY = 1214000L; // 初始 Y 位置
                    const long verticalSpacing = 300000; // 间距 100,000 EMU

                    //string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ±2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50Ω\nChip Size: 1.766 x 2.0 x 0.05mm";
                    //string x = "Feautures\n" + table.GetFeatureTableInfo();

                    // 写入feature
                    long height1 = PptModifier.AddTextBox(slide, pptDataModel.FirstPage.FeaturesText.TrimEnd(), 914400, currentY);
                    // 添加文本框并获取实际高度
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
                    PptModifier.AddTable6(slide, pptDataModel.FirstPage.ParameterTableData, 914400, currentY, 6100000, 350_0000);
                    currentY += 2000000 + verticalSpacing; // 表格高度 + 间距

                    PptModifier.AddTextBox(slide, pptDataModel.FirstPage.FunctionalBlockDiagramImage.ImageName.TrimEnd(), 914400 + 3_250_000,1314000 + 150_000);
                    PptModifier.AddImage(slidePart, pptDataModel.FirstPage.FunctionalBlockDiagramImage.ImagePath, 914400+ 3200_000, 1314000+ 400_000, 240_0000, 270_0000);


                    #endregion

                    // 曲线是4的倍数， 如果不是， 就表示出错了

                    #region 添加曲线页
                    int curvePageCount;

                    if (pptDataModel.CurvesImagePage.CurveImagesPath.Count % 6 == 0)
                    {
                        // 页数为6发非0倍数
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

                    int originY = textOriginY + 500_000;//第一行
                    int originY2 = textOriginY2 + 500_000;//第二行
                    int originY3 = textOriginY3 + 500_000;//第三行

                    List<(int x, int y)> imgagePositions = new List<(int x, int y)>();// 定位位置， 图像位置抽象为 3行2列的表
                    imgagePositions.Add((originX, originY));
                    imgagePositions.Add((offsetX, originY));

                    imgagePositions.Add((originX, originY2));
                    imgagePositions.Add((offsetX, originY2));

                    imgagePositions.Add((originX, originY3));
                    imgagePositions.Add((offsetX, originY3));


                    List<(int x, int y)> textBoxPositions = new List<(int x, int y)>();// 定位位置， 图像位置抽象为 3行2列的表
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
                    //    // 示例：在新幻灯片上添加表格

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
                    //            // 减去第一个四组
                    //            if (((index - 4) / 4) % 2 == 1) // 4的奇数倍，  奇数倍表示的是 s参数， 偶数倍表示的是nf， psat， oip3， pxdb
                    //            {
                    //                if (index % 4 == 0)//4个s参数， 只需要加一次标题
                    //                {
                    //                    PptModifier.AddTextBoxCenter(newSlide1, text, textBoxPositions.ElementAt(index % 6).x, textBoxPositions.ElementAt(index % 6).y);
                    //                    textIndex++;
                    //                }
                    //            }

                    //            else
                    //            {// 偶数倍，表示的是 nf， psat， oip3， pxdb
                    //             //4个s参数， 只需要加一次标题
                    //                PptModifier.AddTextBoxCenterWH(newSlide1, text, textBoxPositions.ElementAt(index % 6).x, textBoxPositions.ElementAt(index % 6).y, 350_0000, 350_0000);
                    //                textIndex++;
                    //            }



                    //        }

                    //        if (index > 0 && index % 6 == 0)
                    //        {
                    //            index++;
                    //            break;// 6个一组， 所以只需要循环6次， 开启新的一页
                    //        }

                    //    }

                    //}
                   int IMAGES_PER_SLIDE = 6;
                   int IMAGE_TITLE_SKIP_GROUP = 4;

                    {
                        int imageIndex = 0; // 用于跟踪总的图片索引
                        int titleIndex = 0; // 用于跟踪总的标题索引
                        int imagesCount = pptDataModel.CurvesImagePage.CurveImagesPath.Count;
                        for (int i = 0; i < curvePageCount; i++)
                        {
                            var newSlidePart1 = PptModifier.AddNewSlideFromLayout(presentationPart);
                            var newSlide1 = newSlidePart1.Slide;

                            for (int j = 0; j < IMAGES_PER_SLIDE; j++)
                            {
                                // 检查是否已达到图片总数末尾
                                if (imageIndex >= pptDataModel.CurvesImagePage.CurveImagesPath.Count)
                                {
                                    break;
                                }

                                // --- 图片处理 ---
                                string imagePath = pptDataModel.CurvesImagePage.CurveImagesPath.ElementAt(imageIndex);
                                Console.WriteLine(imagePath);

                                PptModifier.AddImage(
                                    newSlidePart1,
                                    imagePath,
                                    imgagePositions.ElementAt(j).x,
                                    imgagePositions.ElementAt(j).y,
                                    2_500_000, 2_000_000
                                );

                                // --- 标题处理局部变量初始化 ---
                                string titleText = "";
                                bool shouldAddTitle = false;

                                // 默认值，以防止未进入 else 块时使用
                                int adjustedIndex = -1;
                                bool isSParameterGroup = false;

                                // 检查标题索引是否越界（必须在 ElementAt() 之前）
                                if (titleIndex >= pptDataModel.CurvesImagePage.CurveTitles.Count)
                                {
                                    imageIndex++;
                                    continue; // 没有更多标题了
                                }
                                titleText = pptDataModel.CurvesImagePage.CurveTitles.ElementAt(titleIndex);

                                // --- 复杂标题添加逻辑 ---
                                if ((imagesCount/4) % 2 != 0) // 这个时候有单独的那4张图  图片数目为 4 + 8* n  ， 及单独的4张只含有s参量，  后面8个一组表示 还含有 nf, psat, oip3, pxdb
                                {
                                    if (imageIndex == 0)
                                    {
                                        // 第一张图片，必须添加标题
                                        shouldAddTitle = true;
                                        // 注意：这里 j 也是 0
                                    }
                                    else if (imageIndex < IMAGE_TITLE_SKIP_GROUP)
                                    {
                                        // imageIndex > 0 且小于 4 的情况，原始逻辑是 continue (不添加标题)
                                        shouldAddTitle = false;
                                    }
                                    else
                                    {
                                        // 索引大于等于 4 的情况
                                        adjustedIndex = imageIndex - IMAGE_TITLE_SKIP_GROUP;

                                        // 检查调整后的索引是否属于 S 参数组 (4的奇数倍)
                                        // 原始逻辑: ((index - 4) / 4) % 2 == 1
                                        isSParameterGroup = (adjustedIndex / IMAGE_TITLE_SKIP_GROUP) % 2 == 0;

                                        if (isSParameterGroup)
                                        {
                                            // S 参数组：每组 4 张图，只需要加一次标题
                                            if (adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0)
                                            {
                                                shouldAddTitle = true;
                                            }
                                        }
                                        else
                                        {
                                            // 非 S 参数组：每张图都加标题 (偶数倍)
                                            shouldAddTitle = true;
                                        }
                                    }

                                }
                                else
                                {
                                    if (imageIndex == 0)
                                    {
                                        // 第一张图片，必须添加标题
                                        shouldAddTitle = true;
                                        // 注意：这里 j 也是 0
                                    }
                                    else if (imageIndex < IMAGE_TITLE_SKIP_GROUP)
                                    {
                                        // imageIndex > 0 且小于 4 的情况，原始逻辑是 continue (不添加标题)
                                        shouldAddTitle = false;
                                    }
                                    else
                                    {
                                        // 索引大于等于 4 的情况
                                        adjustedIndex = imageIndex - IMAGE_TITLE_SKIP_GROUP;

                                        // 检查调整后的索引是否属于 S 参数组 (4的奇数倍)
                                        // 原始逻辑: ((index - 4) / 4) % 2 == 1
                                        isSParameterGroup = (adjustedIndex / IMAGE_TITLE_SKIP_GROUP) % 2 == 1;

                                        if (isSParameterGroup)
                                        {
                                            // S 参数组：每组 4 张图，只需要加一次标题
                                            if (adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0)
                                            {
                                                shouldAddTitle = true;
                                            }
                                        }
                                        else
                                        {
                                            // 非 S 参数组：每张图都加标题 (偶数倍)
                                            shouldAddTitle = true;
                                        }
                                    }

                                }



                                //if (imageIndex == 0)
                                //        {
                                //            // 第一张图片，必须添加标题
                                //            shouldAddTitle = true;
                                //            // 注意：这里 j 也是 0
                                //        }
                                //        else if (imageIndex < IMAGE_TITLE_SKIP_GROUP)
                                //        {
                                //            // imageIndex > 0 且小于 4 的情况，原始逻辑是 continue (不添加标题)
                                //            shouldAddTitle = false;
                                //        }
                                //        else
                                //        {
                                //            // 索引大于等于 4 的情况
                                //            adjustedIndex = imageIndex - IMAGE_TITLE_SKIP_GROUP;

                                //            // 检查调整后的索引是否属于 S 参数组 (4的奇数倍)
                                //            // 原始逻辑: ((index - 4) / 4) % 2 == 1
                                //            isSParameterGroup = (adjustedIndex / IMAGE_TITLE_SKIP_GROUP) % 2 == 1;

                                //            if (isSParameterGroup)
                                //            {
                                //                // S 参数组：每组 4 张图，只需要加一次标题
                                //                if (adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0)
                                //                {
                                //                    shouldAddTitle = true;
                                //                }
                                //            }
                                //            else
                                //            {
                                //                // 非 S 参数组：每张图都加标题 (偶数倍)
                                //                shouldAddTitle = true;
                                //            }
                                //        }



                                // --- 执行标题添加 ---
                                if (shouldAddTitle)
                                {
                                    // 这里的 if 条件需要注意：只有当 isSParameterGroup 是 true 
                                    // 且 adjustedIndex % 4 == 0 (即 S 参数组的第一个图) 时，才使用 AddTextBoxCenter

                                    // imageIndex == 0 的情况 isSParameterGroup 和 adjustedIndex 都是默认值，
                                    // 需要确保它不进入 AddTextBoxCenter 的条件

                                    // 优化：使用一个更明确的条件来区分调用方法
                                    bool useAddTextBoxCenter = (imageIndex == 0) ||
                                                               (isSParameterGroup && adjustedIndex % IMAGE_TITLE_SKIP_GROUP == 0);


                                    if (useAddTextBoxCenter)
                                    {
                                        // 用于 imageIndex=0 的情况 和 S 参数组只添加一次的情况
                                        PptModifier.AddTextBoxCenter(
                                            newSlide1,
                                            titleText,
                                            textBoxPositions.ElementAt(j).x,
                                            textBoxPositions.ElementAt(j).y,
                                            1200,500_0000
                                        );
                                    }
                                    else
                                    {
                                        // 用于非 S 参数组的情况
                                        PptModifier.AddTextBoxCenterWH(
                                            newSlide1,
                                            titleText,
                                            textBoxPositions.ElementAt(j).x,
                                            textBoxPositions.ElementAt(j).y,
                                            240_0000, 350_0000
                                        );
                                    }
                                    titleIndex++; // 只有在添加了标题后才增加标题索引
                                }


                                imageIndex++; // 无论是否添加标题，图片索引都要递增

                                // --- 原始的分页判断 ---
                                if (imageIndex > 0 && imageIndex % IMAGES_PER_SLIDE == 0)
                                {
                                    // 已经处理完 6 张图，准备进入下一页
                                    break;
                                }
                            } // 结束当前幻灯片的 6 张图片循环
                        } // 结束分页循环
                    }
                    #endregion





                    #region 倒数第5页
                    var newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    var newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    //info = "Absolute Maximum Ratings";// \u2103 是摄氏度的符号
                    PptModifier.AddTextBoxUnderline(newSlide, pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTableTitle, 814400, originY, 3500_000, 300000);

                    //info = "Typical Supply Current vs. VD,VG";// \u2103 是摄氏度的符号
                    var height = PptModifier.AddTextBoxUnderline(newSlide, pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTableTitle, 914400 + 2700000 + 700000, originY, 3500_000, 300000);
                    currentY += height + 50000;
                    PptModifier.AddTable7(newSlide, pptDataModel.EndToFront5Page.AbsoluteMaximumRatingsTable, 814400, currentY, 330_0000, 3000000);
                    PptModifier.AddTableAverageWidth(newSlide, pptDataModel.EndToFront5Page.TypicalSupplyCurrentVgTable, 101_4400 + 2700000 + 700000, currentY, 2300000, 110_0000);
                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\2.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.EndToFront5Page.WarningImage.ImagePath, 914400 + 2700000 + 600000, currentY + 220_0000, 50_0000, 50_0000);
                    //info = "ELECTROSTATIC SENSITIVE DEVICE\n OBSERVE HANDLING PRECAUTIONS";
                    PptModifier.AddTextBox3(newSlide, pptDataModel.EndToFront5Page.WarningText, 91_4400 + 3000000 + 600000 + 200_000, currentY + 220_0000);

                    #endregion

                    #region 倒数第4页
                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    //info = "Outline Drawing: \nAll Dimensions in μm";
                    //height = PptModifier.AddTextBoxCenter2(newSlide, pptDataModel.EndToFront4Page.PinImage.ImageName, originX, currentY);
                    height = PptModifier.AddTextBoxCenter2(newSlide, "Outline Drawing: \nAll Dimensions in μm", originX, currentY);

                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\3.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.EndToFront4Page.PinImage.ImagePath, 1014400, currentY + 450_000, 500_0000, 500_0000);

                    //info = "Notes:\n1. Die thickness: 50μm\n2. VD bond pad is 75*75μm2 \n3. VG bond pad is 75*75μm2 \n4. RF IN/OUT bond pad is 50*86μm2 \n5. Bond pad metalization: Gold\n6. Backside metalization: Gold\n";
                    PptModifier.AddTextBox2(newSlide, pptDataModel.EndToFront4Page.NoteText, 914400, currentY + 800_000 + 500_0000 + 100_000);



                    #endregion


                    #region 倒数第3页
                    newSlidePart = PptModifier.AddNewSlideFromLayout(presentationPart);
                    newSlide = newSlidePart.Slide;
                    originX = 914400;
                    originY = 1314000;
                    currentY = originY;
                    //info = "Assembly Drawing";
                    height = PptModifier.AddTextBoxCenter(newSlide, pptDataModel.EndToFront3Page.StructImage.ImageName, originX, currentY);

                    //pic1 = @"F:\PROJECT\ChipManualGeneration\exe\4.png";
                    PptModifier.AddImage(newSlidePart, pptDataModel.EndToFront3Page.StructImage.ImagePath, 1014400, currentY + height + 100, 550_0000, 350_0000);

                    PptModifier.AddTable4(newSlide, pptDataModel.EndToFront3Page.Description, 914400, currentY + height + 100 + 350_0000 + 50_000, 320_0000, 200_0000);

                    PptModifier.AddTable4(newSlide, pptDataModel.EndToFront3Page.Description2, 914400, currentY + height + 100 + 350_0000 + 1000 + 200_0000 + 150_000, 600_0000, 200_0000);

                    #endregion




                    #region  倒数第二页
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
                    info = "Direct Mounting\n1.Typically, the die is mounted directly on the ground plane.\n2.If the thickness difference between the?substrate (thickness?c)?and the?die (thickness?d)?exceeds?0.05 mm (i.e.,?c?–?d?> 0.05 mm), it is recommended to first mount the die on a?heat spreader, then attach the heat spreader to the ground plane.\r\n3.Heat Spreader Material: Molybdenum-copper (MoCu) alloy is commonly used.\r\n4.Heat Sink Thickness (b): Should be within the range of?(c?–?d?– 0.05 mm)?to?(c?–?d?+ 0.05 mm).\r\n5.Spacing (a): The gap between the bare die and the 50Ω transmission line should typically be?0.05 mm to 0.1 mm.\r\nIf the application frequency is higher than 40GHz, then this gap is recommended to be 0.05mm\r\nWire Bonding Interconnection\r\nThe connection between the die and the 50Ω transmission line is usually made using?25 ?m diameter gold (Au) wires, bonded via?wedge bonding?or?ball bonding?processes.\r\nDie Attachment Methods\r\n1.Conductive Epoxy:\r\nAfter adhesive application, cure according to the manufacturer’s recommended temperature profile.\r\n2.Au-Sn80/20 Eutectic Bonding:\r\nUse preformed?Au-Sn80/20 solder preforms.\r\nPerform bonding in an inert atmosphere (N? or forming gas: 90% N? + 10% H?).\r\nKeep the time above?320°C?to?less than 20 seconds?to prevent excessive intermetallic formation.\r\n";
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


     


        private void HandleQueryFinished(object sender, EventArgs e)
        {
            Console.WriteLine("Logger: Query has finished.");
            
            ////“Application”是“Microsoft.Office.Interop.PowerPoint.Application”和“System.Windows.Application”之间的不明确的引用
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                contentGridRight.Visibility = Visibility.Visible;
                test2();  // 在主线程执行 UI 操作
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



        private void HandleHomeBtnClick(object sender, EventArgs e)
        {
            HiddenAll();
            taskMangeControls.RefreshTask();
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
                    //separator.Visibility = Visibility.Visible;
                    contentGrid.Visibility = Visibility.Visible;
                    return;
                }
            }
            if (Global.TaskModel != null && vm.ContentTitle != null)
            {
                treeView.Visibility = Visibility.Visible;
                //separator.Visibility = Visibility.Visible;
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

        private async void HandleTaskExcute(object sender, TaskTableItem task)
        {
            //var sqlServer = new TaskRepository();
            //var existingModel = await sqlServer.GetOperationByIdAsync(Convert.ToInt32(task.ID));
            //Global.OperationModel = existingModel;
            //if (existingModel != null)
            //{
            //    if (existingModel.DataReady)
            //    {

            //        var filterCond = new FileterConditionModel();
            //        filterCond.PN = existingModel.PN;
            //        filterCond.ON = existingModel.SN;
            //        filterCond.StartDateTime = existingModel.StartDateTime;
            //        filterCond.StopDateTime = existingModel.EndDateTime;

            //        string[] conditions = existingModel.Condition.Trim(',').Split(',');
            //        string[] vd_vgs = conditions.ElementAt(0).Split(';');

            //        foreach (var item in vd_vgs)
            //        {
            //            filterCond.VD_VG_Conditon.Add(item.Trim(';'));
            //        }
            //        filterCond.Min = Convert.ToDouble(conditions.ElementAt(1));
            //        filterCond.Max = Convert.ToDouble(conditions.ElementAt(2));
            //        int count = conditions.Count();
            //        for (int i = 3; i < count; i++)
            //        {
            //            filterCond.FreqBands.Add(conditions.ElementAt(i));

            //        }

            //        //await Task.Run(() =>
            //        //{
            //        //    filter.SetFileterCondition(filterCond);
            //        //});

            //        filter.SetFileterCondition(filterCond);
            //        await Task.Run(() =>
            //        {
            //            filter.Btn_Next_Clicked(null, null);
            //        });
            //        //filter.Btn_Next_Clicked(null, null);

            //        //filter.Btn_Calcute_Click(null, null);
            //        await Task.Run(() =>
            //        {
            //            filter.Btn_Calcute_Click(null, null);
            //        });
            //        vm.ContentTitle = Global.TaskModel.TaskName +"-" + "Amplifier" +"-" + "MML806";


            //    }


            //}
            //_task = task;
            //HandleOperationBtnClick(null, null);

            SetBtnVisibility(false);
            HiddenAll();
            curves.Clear();
            _OperationPage.Visibility = Visibility.Visible;

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
            //separator.Visibility = Visibility.Collapsed;
            welComeStackPanel.Visibility = Visibility.Collapsed;
            home.Visibility = Visibility.Collapsed;
            contentGrid.Visibility = Visibility.Collapsed;
            log.Visibility = Visibility.Collapsed;
            addPage.Visibility = Visibility.Collapsed;
            newTaskPage.Visibility = Visibility.Collapsed;
            _OperationPage.Visibility = Visibility.Collapsed;
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

        private void Btn_ViewLog_Click(object sender, RoutedEventArgs e)
        {
            OperationWindowHiddenAll();
        }

        private async void Btn_ViewCurves_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                SetBtnVisibility(true);
                OperationWindowHiddenAll();
                curves.Visibility = Visibility.Visible;
                curves.Clear();
                _plotNames.Clear();
                vm.IsBusy = true;
                vm.BusyMessage = "Loading curves...";
                //var sqlServer = new TaskSqlServerRepository();
                //var taskItems = await sqlServer.GetAllTasksAsync();
                var taskItem = Global.TaskModel;
                //Global.TaskModel
                #region 在将之前的操作条件 反序列化，得到结果
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                };
                //var conditons = JsonSerializer.Deserialize<TaskFrequencyConfig>(Global.TaskModel.Conditions, options);

                var conditons = JsonSerializer.Deserialize<TaskFrequencyConfig>(taskItem.Conditions, options);

                #endregion


                //从FTP服务器下载数据
                //string ftpRemotePath = Global.TaskModel.Major +"\\" + Global.TaskModel.Minor + "\\" + Global.TaskModel.TaskName;
                string ftpRemotePath = taskItem.Major + "\\" + taskItem.Minor + "\\" + taskItem.TaskName;

                if (await FtpClient.DownloadFolderAsync(ftpRemotePath, Global.TempBasePath))
                {
                    #region 将文件夹的数据进行分裂
                    var finder = new TextFileFinder(
                                   rootDirectory: Global.TempBasePath,
                                    extensions: new[] { ".txt", "s2p" }
                                   );
                    //var finder = new TextFileFinder(); // 或你的文件查找器
                    var allFiles = finder.FindAllTextFiles(); // 返回相对路径列表

                    // 将里面的文件处理并分组
                    var filesByGroup = AmplifierFileProcessor.ProcessFiles(allFiles);
                    #endregion

                    var fileterCondi = new FileterConditionModel
                    {
                        Min = Convert.ToDouble(conditons.MinFrequency),
                        Max = Convert.ToDouble(conditons.MaxFrequency),

                        VD_VG_Conditon = conditons.ParameterItems,
                    };
                    if (conditons.Band1 != null && conditons.Band1 != "0 - 0")// 这个"0 - 0" 是因为有些band没有用到时初始化成这个特殊字符
                        fileterCondi.FreqBands.Add(conditons.Band1);
                    if (conditons.Band2 != null&&conditons.Band2 != "0 - 0")
                        fileterCondi.FreqBands.Add(conditons.Band2);
                    if (conditons.Band3 != null && conditons.Band3 != "0 - 0")
                        fileterCondi.FreqBands.Add(conditons.Band3);



                    _filterCondition = fileterCondi;
                    //if (fileterCondi.VD_VG_Conditon.Count > 1)
                    //{

                    //    GeneratePlots(filesByGroup, fileterCondi);
                    //    var temp = "Measurement Plots: S-parameters\n" + fileterCondi.VD_VG_Conditon.ElementAt(0).Replace('&', ',');
                    //    _plotNames.Add(temp);
                    //}
                    if (fileterCondi.VD_VG_Conditon.Count > 1)
                    {

                        GeneratePlots(filesByGroup, fileterCondi, conditons.SelectedEntry);
                        var temp = "Measurement Plots: S-parameters\n" + fileterCondi.VD_VG_Conditon.ElementAt(0).Replace('&', ',');
                        _plotNames.Add(temp);
                    }

                    Temperature25FilePathModel filesModel25 = new Temperature25FilePathModel();
                    const string tempKey = "25.0deg";
                    var conditions = fileterCondi.VD_VG_Conditon;

                    // 1. 定义所有要处理的属性及其对应的目标列表
                    // 使用元组数组来映射来源字典和目标列表
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
                    // 2. 遍历所有映射并统一处理
                    foreach (var mapping in fileMappings)
                    {
                        // 尝试从当前来源字典中获取文件列表
                        if (mapping.sourceDict.TryGetValue(tempKey, out var files))
                        {
                            // 使用通用的筛选函数获取匹配的文件
                            var matchingFiles = GetMatchingFiles(files, conditions);

                            // 将结果一次性添加到目标列表中
                            mapping.targetList.AddRange(matchingFiles);
                        }
                    }


                    GeneratePlotsByTemperaturen(filesModel25, fileterCondi.Min, fileterCondi.Max, LegendTextType.Elec);

                    #region 三温图

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



                    foreach (var subCon in fileterCondi.VD_VG_Conditon)
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
                        actually.SList = SelectUnique5VFilePerTemperature(filesModelvg.SList, vd);
                        actually.PxdbList = SelectUnique5VFilePerTemperature(filesModelvg.PxdbList, vd);
                        actually.OIP3List = SelectUnique5VFilePerTemperature(filesModelvg.OIP3List, vd);
                        actually.PsatList = SelectUnique5VFilePerTemperature(filesModelvg.PsatList, vd);
                        actually.NFList = SelectUnique5VFilePerTemperature(filesModelvg.NFList, vd);
                        filesModelVgs.Add(actually);
                    }
                    foreach (var item in filesModelVgs)// 不同的VD下， 不同的三温度图
                    {
                        if (item.SList.Count > 0)
                        {
                            List<string> culFiles = new List<string>();
                            culFiles.Add(item.SList.ElementAt(0));
                            culFiles.Add(item.PxdbList.ElementAt(0));
                            culFiles.Add(item.OIP3List.ElementAt(0));
                            culFiles.Add(item.PsatList.ElementAt(0));
                            culFiles.Add(item.NFList.ElementAt(0));
                            CalcuteParameter(culFiles, fileterCondi.FreqBands);

                            break;
                        }
                    }


                    for (int i = 0; i < filesModelVgs.Count; i++)// 不同的VD下， 不同的三温度图
                    {
                        if (filesModelVgs.ElementAt(i).SList.Count > 0)
                        {
                            GeneratePlotsByTemperaturen(filesModelVgs.ElementAt(i), fileterCondi.Min, fileterCondi.Max, LegendTextType.Temp);
                            // 这个数组用于处理参数列表， 是用来计算的

                            var tem11 = "Measurement Plots: S-parameters\n" + fileterCondi.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                            var tem22 = "Measurement Plots: P1dB\n" + fileterCondi.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                            var tem33 = "Measurement Plots: OIP3\n" + fileterCondi.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                            var tem44 = "Measurement Plots: Psat\n" + fileterCondi.VD_VG_Conditon.ElementAt(i).Replace('&', ',');
                            var tem55 = "Measurement Plots: Noise Figure\n" + fileterCondi.VD_VG_Conditon.ElementAt(i).Replace('&', ',');

                            _plotNames.Add(tem11);
                            _plotNames.Add(tem22);
                            _plotNames.Add(tem33);
                            _plotNames.Add(tem44);
                            _plotNames.Add(tem55);
                            //CalcuteParameter(List<string> files, List<string> freqBands)
                        }


                    }


                }
                else
                {
                    MessageBox.Show("Download failed!");
                }
                #endregion

                #region 加载表格
                string jsonPath = System.IO.Path.Combine(Global.TempBasePath, "table.json");
                if (taskItem.TableUpdate == false)
                {
                    #region 先创建表格
                    //if (File.Exists(jsonPath))
                    //{
                    //    Console.WriteLine($"{jsonPath} not exist");
                    //    return;
                    //}
                    SaveConfigs();// 先创建一个基本的productdata 的json文件


                    #endregion
                    #region 再更新表格

                    var productData = LoadTableInfo(jsonPath);
                    if (productData?.Tables == null)
                    {
                        Console.WriteLine("数据加载失败或 Tables 字段为空。");
                        return;
                    }
                    string newFrequencyValue = _filterCondition.Min.ToString() + "-" + _filterCondition.Max.ToString() + "GHz";

                    var basicFreqItem = productData.Tables.BasicParameters
                                         .FirstOrDefault(item => item.Key == "{Frequency Range}");

                    if (basicFreqItem != null)
                    {

                        basicFreqItem.Value = newFrequencyValue;
                        Console.WriteLine($"BasicParameters 中的频率已修改为: {basicFreqItem.Value}");
                    }
                    else
                    {
                        Console.WriteLine("未找到 BasicParameters 中的 {Frequency Range} 项。");
                    }
                    var rightBarItem = productData.Tables.BasicParameters
                                       .FirstOrDefault(item => item.Key == "{right bar info}");
                   
                    if (rightBarItem != null)
                    {
                        // 正则表达式模式解释：
                        // \d+      : 匹配一个或多个数字 (如 45)
                        // [\s–-]+  : 匹配一个或多个空格、长破折号 (–) 或短破折号 (-)
                        // \d+GHz   : 匹配一个或多个数字后跟 GHz
                        string pattern = @"\d+[\s–-]+\d+GHz";
                        string updatedString = Regex.Replace(rightBarItem.Value, pattern, newFrequencyValue);
                        rightBarItem.Value = updatedString;
                    }
                    else
                    { 
                        Console.WriteLine("未找到 BasicParameters 中的 {right bar info} 项。");
                    }

                    var chipItem = productData.Tables.BasicParameters
                                      .FirstOrDefault(item => item.Key == "{Top PN}");
                    if (chipItem != null)
                    {
                        chipItem.Value = taskItem.TaskName;
                    }
                    else 
                    {
                    
                    }
                

                     var featureFreqItem = productData.Tables.FeatureParameters
                                                .FirstOrDefault(item => item.Key == "Frequency");
                    if (featureFreqItem != null)
                    {

                        featureFreqItem.Value = newFrequencyValue;
                        Console.WriteLine($"FeatureParameters 中的频率已修改为: {featureFreqItem.Value}");
                    }
                    else
                    {
                        Console.WriteLine("未找到 FeatureParameters 中的 Frequency 项。");
                    }
                    var serializeOptions = new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        // 同样添加编码设置以避免不必要的转义
                        // = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    };
                    string updatedJsonString = JsonSerializer.Serialize(productData, serializeOptions);
                    File.WriteAllText(jsonPath, updatedJsonString);
                    #endregion

                }
                else
                {
                    #region 从ftp服务器上下载最新json 文件
                    string remotePath = System.IO.Path.Combine(Global.FtpRootPath, "table.json");
                    string localPath = System.IO.Path.Combine(Global.TempBasePath, "table.json");
                    await FtpClient.DownloadFileAsync(remotePath, localPath);
                    #endregion

                }
                string filePath = System.IO.Path.Combine(Global.TempBasePath, "table.json");
                string jsonString = File.ReadAllText(filePath);
                table.LoadProductDataToViewModel(jsonString);
                #endregion


            }
            catch (Exception ex) {

                MessageBox.Show(ex.Message, "",MessageBoxButton.OK,MessageBoxImage.Error);
            }

            finally
            {
                vm.IsBusy = false;

            }
        }
        private ProductData LoadTableInfo(string jsonPath)
        {
            ProductData resoult = null;
            try
            {
                if (File.Exists(jsonPath))
                {
                    string jsonString = File.ReadAllText(jsonPath);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                    };
                    resoult = JsonSerializer.Deserialize<ProductData>(jsonString, options);
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                
            }


            return resoult;

        }

        bool WriteTableInfo(string jsonPath, ProductData productData)
        {
            bool resoult = false;
            try
            {

            }
            catch
            {

            }
            
            return resoult;
        }
        private void Btn_ViewWorld_Click(object sender, RoutedEventArgs e)
        {
            OperationWindowHiddenAll();
            table.Visibility = Visibility.Visible;
        }
        private void Btn_ViewImage_Click(object sender, RoutedEventArgs e)
        {
            OperationWindowHiddenAll();
            images.Visibility = Visibility.Visible;
        }
        private void Btn_OPBack_Click(object sender, RoutedEventArgs e)
        {
            HandleHomeBtnClick(null, null);

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

                //string pptFile = @"resources\files\demo.pptx"; 用这个方式执行转换会转换失败
                //string pdfFile = @"resources\files\demo.pdf";
                // 在后台线程执行（避免 UI 冻结）
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

        private async void Btn_Preview_PPT_Model_Clicked(object sender, RoutedEventArgs e)
        {
            //string pdfFile = "";
            //if (vm.ContentTitle == "Amplifier--MM809")
            //    pdfFile = @"F:\PROJECT\ChipManualGeneration\放大器\MML806_V3.pdf";
            //else if (vm.ContentTitle == "Amplifier--MM808")
            //    pdfFile = "F:\\PROJECT\\ChipManualGeneration\\放大器\\MML814_V3.0.1.pdf";
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

        private async void Btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //string filePath = System.IO.Path.Combine(Global.FileBasePath, "table.json");
                //string jsonString = File.ReadAllText(filePath);
                //table.LoadProductDataToViewModel(jsonString);

                string fileName = "";
                var basicTable = table.GetBasicTableData();
                string[] arry = basicTable.ElementAt(1).info.Split('.');
                //fileName = basicTable.ElementAt(0).info+"_" + basicTable.ElementAt(1).info.Substring(0, 2)  + ".pptx";
                fileName = basicTable.ElementAt(0).info + "_" + arry[0] + ".pptx";
                var newRepository = new TaskSqlServerRepository();
                var newItem = await newRepository.GetTaskByIdAsync(Global.TaskModel.ID);
                newItem.PptName = fileName;
                await newRepository.UpdateTaskAsync(newItem); 



                var dialog = new SaveFileDialog
                    {
                        Title = "Save As",
                        Filter = "PPT File (*.pptx)|*.pptx",
                        InitialDirectory = Global.AppBaseUrl,
                        FileName = fileName,
                    };
                        if (dialog.ShowDialog() == true)
                        {
                            string pptFile = dialog.FileName;
                             _fileFolerPath = pptFile.Replace(System.IO.Path.GetFileName(pptFile), "");
                             SaveConfigs();
                            vm.IsBusy = true;
                            vm.BusyMessage = "Generating PPT...";
                            string appDir = AppDomain.CurrentDomain.BaseDirectory;
                            string picsPath = System.IO.Path.Combine(appDir, "resources", "pic");
                            System.IO.Directory.Delete(picsPath, true);
                            System.IO.Directory.CreateDirectory(picsPath);
                            curves.SaveAllPlot(picsPath);
                            PptDataModeFactory();
                            //string pptFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pptx");
                            string pdfFile = System.IO.Path.Combine(appDir, "resources", "files", "demo.pdf");
                            await Task.Run(() => GeneratePPT(pptFile));
                            await Task.Run(() => PptToPdfConverter.Convert(pptFile, pdfFile));
                            MessageBox.Show("Generate PPT success!","", MessageBoxButton.OK,MessageBoxImage.Information);
                        }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Tips", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            vm.IsBusy = false;

        }

        private async void Btn_Save_click(object sender, RoutedEventArgs e)
        {
            SaveConfigs();

            
            string remotePath = System.IO.Path.Combine(Global.FtpRootPath,"table.json" );
            string localPath = System.IO.Path.Combine(Global.TempBasePath, "table.json");
            if (await FtpClient.UploadFileAsync(localPath, remotePath))
            {
                var newRepository = new TaskSqlServerRepository();
                var newItem = await newRepository.GetTaskByIdAsync(Global.TaskModel.ID);
                newItem.TableUpdate = true;
                
                newItem.EndDate = DateTime.Now;
                //newItem.FilesStatus = true;
                await newRepository.UpdateTaskAsync(newItem);
                MessageBox.Show("Upload success!");
            }
            else
            {
                MessageBox.Show("Upload failed!");
            }


        }
        private async void Btn_Upload_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select a file",
                Filter = "PowerPoint File (*.pptx)|*.pptx",
                InitialDirectory = _fileFolerPath
            };

            if (dialog.ShowDialog() == true)
            {
                SaveConfigs();
                string filePath = dialog.FileName;
                string remotePath = System.IO.Path.Combine(Global.FtpRootPath, System.IO.Path.GetFileName(filePath));
                if (await FtpClient.UploadFileAsync(filePath, remotePath))
                {
                    var newRepository = new TaskSqlServerRepository();
                    var newItem = await newRepository.GetTaskByIdAsync(Global.TaskModel.ID);
                    newItem.Status = "Completed";
                    newItem.EndDate = DateTime.Now;
                    newItem.FilesStatus=true;
                    await newRepository.UpdateTaskAsync(newItem);
                    MessageBox.Show("Upload success!");
                }
                else
                {
                    MessageBox.Show("Upload failed!");
                }


            }
        }


        public void SetBtnVisibility(bool enable)
        {
            _btnVewWord.IsEnabled = enable;
            _btnSave.IsEnabled = enable;
            _btnGenPPT.IsEnabled = enable;
        }


        private void OperationWindowHiddenAll()
        {
            table.Visibility = Visibility.Collapsed;
            curves.Visibility = Visibility.Collapsed;
            table.Visibility =Visibility.Collapsed;
            images.Visibility = Visibility.Collapsed;
        }
        private void SaveConfigs() 
        {
            var tableData = table.CreateProductDataFromViewModel();
            string filePath = System.IO.Path.Combine(Global.TempBasePath, "table.json");
            WriteProductJsonFile(tableData, filePath);

        }

        public  void WriteProductJsonFile(ProductData data, string filePath)
        {
            try
            {

                //var options = new JsonSerializerOptions
                //{
                //    WriteIndented = true
                //};


                //string jsonString = JsonSerializer.Serialize(data, options);

                //File.WriteAllText(filePath, jsonString);
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                
                string jsonString = JsonSerializer.Serialize(data, options);

              
                // File.CreateText(filePath) 创建一个 StreamWriter， StreamWriter 实现了 IDisposable
                // 离开 using 块时，会自动调用 writer.Dispose()，从而关闭底层文件流。
                using (StreamWriter writer = File.CreateText(filePath))
                {
                    writer.Write(jsonString);
                    // 不需要手动调用 Flush()，因为 StreamWriter.Dispose() 会自动调用 Flush()。
                } 


            }
            catch (Exception ex)
            {
                Console.WriteLine($"序列化错误: {ex.Message}");
            }
        }

        public ProductData ReadJsonFile(string filePath)
        {
            try { 
                //读取时 没有使用异步 无需使用using
                string jsonString = File.ReadAllText(filePath);
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                ProductData data = JsonSerializer.Deserialize<ProductData>(jsonString, options);
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"反序列化错误: {ex.Message}");
                return null;
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

        public ObservableCollection<DeviceTreeViewItemModel> TreeViewSources { get; set; } = new ObservableCollection<DeviceTreeViewItemModel>();

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
            // 如果 Children 集合为空（或 Count 为 0），则为叶子节点
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
                // 当子集合变化时，通知绑定系统 IsLeaf 属性可能已改变
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


    public class DeviceTreeViewItemModelSigle : ObservableObject
    {


        private string _header;
        public string Header
        {
            get => _header;
            set => SetProperty(ref _header, value); // 假设 SetProperty 存在于基类
        }



        // XAML 中绑定了 IsChecked="{Binding IsSelectedInModel, Mode=TwoWay}"
        private bool _isSelectedInModel;
        public bool IsSelectedInModel
        {
            get => _isSelectedInModel;
            set
            {
                if (SetProperty(ref _isSelectedInModel, value))
                {

                    UpdateChildrenSelection(value);
                }

            }
        }
        public void UpdateChildrenSelection(bool isSelected)
        {

            if (Children.Count > 0)
            {
                foreach (var child in Children)
                {

                    child.IsSelectedInModel = isSelected;
                }
            }
        }

        public ObservableCollection<DeviceTreeViewItemModel> Children { get; } =
            new ObservableCollection<DeviceTreeViewItemModel>();



        public DeviceTreeViewItemModelSigle(string header)
        {
            Header = header;
        }

        public override string ToString()
        {
            return Header;
        }




    }

}

