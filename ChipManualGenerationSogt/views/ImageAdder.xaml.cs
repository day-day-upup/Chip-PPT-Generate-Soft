using DocumentFormat.OpenXml.Office2010.CustomUI;
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
    /// ImageAdder.xaml 的交互逻辑
    /// </summary>
    public partial class ImageAdder : UserControl
    {
        // 存储图片路径的字典
        private Dictionary<string, (string topImage, string bottomImage)> _imageMappings;
        //private Dictionary<string, string > _imageMappings;
        string _currentImageName;
        public ImageAdder()
        {
            InitializeComponent();
            InitializeImageMappings();
            // 默认选择第一项
            ImageListBox.SelectedIndex = 0;
        }
        private void InitializeImageMappings()
        {
            string BasePath = "F:\\PROJECT\\ChipManualGeneration\\exe\\app\\ChipManualGenerationSogt\\bin\\Debug\\resources\\files";
            _imageMappings = new Dictionary<string, (string, string)>
            {
                //["Functional Block Diagram"] = (
                //    "F:\\PROJECT\\ChipManualGeneration\\exe\\1.png",
                //    ""
                    
                //),
                //["Outline Drawing"] = (
                //    "F:\\PROJECT\\ChipManualGeneration\\exe\\3.png",
                //     ""
                //),
                //["Assembly Drawing"] = (
                //    "F:\\PROJECT\\ChipManualGeneration\\exe\\4.png",
                //     ""
                //),
                //["Biasing and Operation"] = (
                //    "F:\\PROJECT\\ChipManualGeneration\\exe\\5.png",
                //    ""
                //),
                //["Mounting Bonding Techniques for MMICs"] = (
                //    "F:\\PROJECT\\ChipManualGeneration\\exe\\6.png",
                //     ""
                //)

                 ["Functional Block Diagram"] = (
                    System.IO.Path.Combine(BasePath, "功能图.png"),
                    ""

                ),
                ["Outline Drawing"] = (
                    System.IO.Path.Combine(BasePath, "外形图.png"),
                     ""
                ),
                ["Assembly Drawing"] = (
                    System.IO.Path.Combine(BasePath, "装配图.png"),
                     ""
                ),
                ["Biasing and Operation"] = (
                    System.IO.Path.Combine(BasePath, "框图.png"),
                    ""
                ),
                ["Mounting & Bonding Techniques for MMICs"] = (
                     System.IO.Path.Combine(BasePath, "芯片安装图.png"),
                     ""
                )
            };

        }
        private void ChangeImage_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "图像文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.ico|所有文件|*.*",
                Title = "请选择一张图片",
                Multiselect = false // 只允许选择一个文件
            };

            // 显示对话框并检查用户是否点击了“确定”
            bool? result = openFileDialog.ShowDialog();

            if (result == true) // 用户选择了文件并点击了“打开”
            {
                string selectedFilePath = openFileDialog.FileName;
                (string newFilePath, string otherValue) newTuple = (selectedFilePath, ""); // 假设 Item2 被命名为 otherValue
                
                //try
                //{
                //    // 创建 BitmapImage 并加载文件
                //    BitmapImage bitmap = new BitmapImage();
                //    bitmap.BeginInit();
                //    bitmap.UriSource = new Uri(selectedFilePath);
                //    bitmap.CacheOption = BitmapCacheOption.OnLoad; // 关键：加载后释放文件句柄
                //    bitmap.EndInit();
                //    bitmap.Freeze(); // 可选：提升性能，使图像可跨线程使用

                //    // 更新 Image 控件的 Source
                //    image.Source = bitmap;
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show($"无法加载图片：{ex.Message}", "错误",
                //        MessageBoxButton.OK, MessageBoxImage.Error);
                //}
                if (LoadImage(TopImage, selectedFilePath))
                {
                    _imageMappings[_currentImageName] = newTuple;
                }


            }

        }

        private void ResizeImage_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ImageListBox?.SelectedItem is ListBoxItem selectedItem)
            {
                string content = selectedItem.Content?.ToString();
                _currentImageName = content;
                if (!string.IsNullOrEmpty(content) && _imageMappings.ContainsKey(content))
                {
                    var (topImagePath, bottomImagePath) = _imageMappings[content];
                    LoadImage(TopImage, topImagePath);
                    LoadImage(BottomImage, bottomImagePath);
                }
            }
        }
        private bool LoadImage(Image imageControl, string imagePath)
        {
            bool resoult = false;
            string actualImagePath = imagePath;
            
            try
            {
                if (string.IsNullOrEmpty(imagePath))
                {
                    imageControl.Source = null;
                    return resoult;
                }

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(imagePath);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();

                imageControl.Source = bitmap;
                resoult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法加载图片 {imagePath}：{ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                imageControl.Source = null;
            }
            return resoult;
        }


        private void LoadImage(Image imageControl, string imagePath, out string newPath)
        {

            string actualImagePath = imagePath;
            newPath = "";

            try
            {
                if (string.IsNullOrEmpty(imagePath))
                {
                    imageControl.Source = null;
                    return;
                }

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(imagePath);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();

                imageControl.Source = bitmap;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法加载图片 {imagePath}：{ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                imageControl.Source = null;
            }
        }

        private void ChangeImage(Image targetImage)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "图像文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.ico|所有文件|*.*",
                Title = "请选择一张图片",
                Multiselect = false
            };

            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                string selectedFilePath = openFileDialog.FileName;

                try
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(selectedFilePath);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    bitmap.Freeze();

                    targetImage.Source = bitmap;

                    // 可选：更新当前选中的图片映射
                    UpdateCurrentImageMapping(targetImage, selectedFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法加载图片：{ex.Message}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UpdateCurrentImageMapping(Image targetImage, string filePath)
        {
            if (ImageListBox?.SelectedItem is ListBoxItem selectedItem)
            {
                string content = selectedItem.Content?.ToString();
                if (!string.IsNullOrEmpty(content) && _imageMappings.ContainsKey(content))
                {
                    var currentMapping = _imageMappings[content];

                    if (targetImage == TopImage)
                    {
                        _imageMappings[content] = (filePath, currentMapping.bottomImage);
                    }
                    else if (targetImage == BottomImage)
                    {
                        _imageMappings[content] = (currentMapping.topImage, filePath);
                    }
                }
            }
        }

        public List<(string name, string filePath)> GetAllImage()
        { 
            var resoult =  new List<(string name, string filePath)>();
            (string name, string filePath) img1 = ("Functional Block Diagram", _imageMappings["Functional Block Diagram"].Item1);
            (string name, string filePath) img2 = ("Outline Drawing", _imageMappings["Outline Drawing"].Item1);
            (string name, string filePath) img3 = ("Assembly Drawing", _imageMappings["Assembly Drawing"].Item1);
            (string name, string filePath) img4 = ("Biasing and Operation", _imageMappings["Biasing and Operation"].Item1);
            (string name, string filePath) img5 = ("Mounting & Bonding Techniques for MMICs", _imageMappings["Mounting & Bonding Techniques for MMICs"].Item1);
            resoult.Add(img1);
            resoult.Add(img2);
            resoult.Add(img3);
            resoult.Add(img4);  
            resoult.Add(img5);

            return resoult;
        }
    }
}
