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
        public ImageAdder()
        {
            InitializeComponent();
            InitializeImageMappings();
            // 默认选择第一项
            ImageListBox.SelectedIndex = 0;
        }
        private void InitializeImageMappings()
        {
            _imageMappings = new Dictionary<string, (string, string)>
            {
                ["Functional Block Diagram"] = (
                    "/resources/pic/FunctionalBlockDiagram.png",
                    "/resources/pic/FunctionalBlockDiagram_Bottom.png"
                ),
                ["Outline Drawing"] = (
                    "/resources/pic/OutlineDrawing.png",
                    "/resources/pic/OutlineDrawing_Bottom.png"
                ),
                ["Assembly Drawing"] = (
                    "/resources/pic/AssemblyDrawing.png",
                    "/resources/pic/AssemblyDrawing_Bottom.png"
                ),
                ["Biasing and Operation"] = (
                    "/resources/pic/BiasingOperation.png",
                    "/resources/pic/BiasingOperation_Bottom.png"
                ),
                ["Mounting Bonding Technigues for MMICs"] = (
                    "/resources/pic/MountingBonding.png",
                    "/resources/pic/MountingBonding_Bottom.png"
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
                LoadImage(TopImage, selectedFilePath);


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
                if (!string.IsNullOrEmpty(content) && _imageMappings.ContainsKey(content))
                {
                    var (topImagePath, bottomImagePath) = _imageMappings[content];
                    LoadImage(TopImage, topImagePath);
                    LoadImage(BottomImage, bottomImagePath);
                }
            }
        }
        private void LoadImage(Image imageControl, string imagePath)
        {
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
    }
}
