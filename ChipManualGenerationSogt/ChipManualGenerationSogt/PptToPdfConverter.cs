using System;
using System.IO;
using System.Runtime.InteropServices;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace ChipManualGenerationSogt
{
    public static class PptToPdfConverter
    {
        public static void Convert(string pptFilePath, string pdfFilePath)
        {
            if (string.IsNullOrWhiteSpace(pptFilePath))
                throw new ArgumentException("PPT file path is null or empty.", nameof(pptFilePath));
            if (string.IsNullOrWhiteSpace(pdfFilePath))
                throw new ArgumentException("PDF file path is null or empty.", nameof(pdfFilePath));

            if (!File.Exists(pptFilePath))
                throw new FileNotFoundException("PPT file not found.", pptFilePath);

            // 确保输出目录存在
            Directory.CreateDirectory(Path.GetDirectoryName(pdfFilePath));

            PPT.Application pptApp = null;
            PPT.Presentation presentation = null;

            try
            {
                pptApp = new PPT.Application();
                // 以静默模式打开（不显示窗口）
                presentation = pptApp.Presentations.Open(
                    FileName: pptFilePath,
                    ReadOnly: Microsoft.Office.Core.MsoTriState.msoTrue,
                    Untitled: Microsoft.Office.Core.MsoTriState.msoFalse,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse
                );

                // 导出为 PDF
                presentation.ExportAsFixedFormat(
                    Path: pdfFilePath,
                    FixedFormatType: PPT.PpFixedFormatType.ppFixedFormatTypePDF,
                    Intent: PPT.PpFixedFormatIntent.ppFixedFormatIntentScreen,
                    FrameSlides: Microsoft.Office.Core.MsoTriState.msoTrue,
                    HandoutOrder: PPT.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    OutputType: PPT.PpPrintOutputType.ppPrintOutputSlides,
                    PrintHiddenSlides: Microsoft.Office.Core.MsoTriState.msoTrue,
                    RangeType: PPT.PpPrintRangeType.ppPrintAll
                );
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Failed to convert PPT to PDF. Ensure PowerPoint is installed.", ex);
            }
            finally
            {
                // 安全释放 COM 对象
                if (presentation != null)
                {
                    presentation.Close();
                    Marshal.ReleaseComObject(presentation);
                }

                if (pptApp != null)
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }

                // 强制垃圾回收（可选，但有助于释放进程）
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
       
    
    }
}


