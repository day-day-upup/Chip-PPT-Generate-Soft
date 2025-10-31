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
    /// ImageAdder.xaml �Ľ����߼�
    /// </summary>
    public partial class ImageAdder : UserControl
    {
        // �洢ͼƬ·�����ֵ�
        private Dictionary<string, (string topImage, string bottomImage)> _imageMappings;
        //private Dictionary<string, string > _imageMappings;
        string _currentImageName;
        public ImageAdder()
        {
            InitializeComponent();
            InitializeImageMappings();
            // Ĭ��ѡ���һ��
            ImageListBox.SelectedIndex = 0;
        }
        private void InitializeImageMappings()
        {
            _imageMappings = new Dictionary<string, (string, string)>
            {
                ["Functional Block Diagram"] = (
                    "F:\\PROJECT\\ChipManualGeneration\\exe\\1.png",
                    "pack://application:,,,/resources/pic/fuction.png"
                    
                ),
                ["Outline Drawing"] = (
                    "F:\\PROJECT\\ChipManualGeneration\\exe\\3.png",
                     "pack://application:,,,/resources/pic/out drawing.png"
                ),
                ["Assembly Drawing"] = (
                    "F:\\PROJECT\\ChipManualGeneration\\exe\\4.png",
                     "pack://application:,,,/resources/pic/a drawing.png"
                ),
                ["Biasing and Operation"] = (
                    "F:\\PROJECT\\ChipManualGeneration\\exe\\5.png",
                    "pack://application:,,,/resources/pic/b drawing.png"
                ),
                ["Mounting Bonding Techniques for MMICs"] = (
                    "F:\\PROJECT\\ChipManualGeneration\\exe\\6.png",
                     "pack://application:,,,/resources/pic/m drawing.png"
                )
            };

        }
        private void ChangeImage_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "ͼ���ļ�|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.ico|�����ļ�|*.*",
                Title = "��ѡ��һ��ͼƬ",
                Multiselect = false // ֻ����ѡ��һ���ļ�
            };

            // ��ʾ�Ի��򲢼���û��Ƿ����ˡ�ȷ����
            bool? result = openFileDialog.ShowDialog();

            if (result == true) // �û�ѡ�����ļ�������ˡ��򿪡�
            {
                string selectedFilePath = openFileDialog.FileName;
                (string newFilePath, string otherValue) newTuple = (selectedFilePath, ""); // ���� Item2 ������Ϊ otherValue
                
                //try
                //{
                //    // ���� BitmapImage �������ļ�
                //    BitmapImage bitmap = new BitmapImage();
                //    bitmap.BeginInit();
                //    bitmap.UriSource = new Uri(selectedFilePath);
                //    bitmap.CacheOption = BitmapCacheOption.OnLoad; // �ؼ������غ��ͷ��ļ����
                //    bitmap.EndInit();
                //    bitmap.Freeze(); // ��ѡ���������ܣ�ʹͼ��ɿ��߳�ʹ��

                //    // ���� Image �ؼ��� Source
                //    image.Source = bitmap;
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show($"�޷�����ͼƬ��{ex.Message}", "����",
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
                MessageBox.Show($"�޷�����ͼƬ {imagePath}��{ex.Message}", "����",
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
                MessageBox.Show($"�޷�����ͼƬ {imagePath}��{ex.Message}", "����",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                imageControl.Source = null;
            }
        }

        private void ChangeImage(Image targetImage)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "ͼ���ļ�|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.ico|�����ļ�|*.*",
                Title = "��ѡ��һ��ͼƬ",
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

                    // ��ѡ�����µ�ǰѡ�е�ͼƬӳ��
                    UpdateCurrentImageMapping(targetImage, selectedFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"�޷�����ͼƬ��{ex.Message}", "����",
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
            (string name, string filePath) img5 = ("Mounting Bonding Techniques for MMICs", _imageMappings["Mounting Bonding Techniques for MMICs"].Item1);
            resoult.Add(img1);
            resoult.Add(img2);
            resoult.Add(img3);
            resoult.Add(img4);  
            resoult.Add(img5);

            return resoult;
        }
    }
}
