using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using ChipManualGenerationSogt.models;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ChipManualGenerationSogt
{
    public class FontConfig
    {
        public int Size { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public string Color { get; set; }
        public string Typeface { get; set; } = "Calibri";
        public A.TextUnderlineValues? Underline { get; set; }
    }
    public class TextBoxConfig
    {
        public string Text { get; set; }
        public FontConfig Font { get; set; }
        public A.TextAlignmentTypeValues Alignment { get; set; } = A.TextAlignmentTypeValues.Left;
        public bool HasBullet { get; set; }
        public long CustomWidth { get; set; }
        public long CustomHeight { get; set; }
        public long OffsetX { get; set; }
        public long OffsetY { get; set; }
    }

    public class SlideMasterTextInfo
    {



    }
    public class PptModifier
    {
        public static string FilePath { get; set; }
        public void InsertTextAndImage(string filePath, string imagePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("未找到PPT文件", filePath);

            var presentationDoc = PresentationDocument.Open(filePath, isEditable: true);
            var presentationPart = presentationDoc.PresentationPart;

            // 获取第一个 slide part（通过关系）
            var slideId = presentationPart.Presentation.SlideIdList?.GetFirstChild<P.SlideId>();
            //var slideId = presentationPart.Presentation.SlideIdList?.FirstOrDefault();
            if (slideId == null)
                throw new InvalidOperationException("PPT中没有幻灯片。");


            // ? 关键：v3.3.0 中用 GetPartById 获取 SlidePart（类型是 OpenXmlPart，但可转为 Slide）
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
            var slide = slidePart.Slide;

            //var currentY = 1214000L; // 初始 Y 位置（1 英寸）
            //const long textBoxHeight = 800000; // 文本框高度（EMU）
            //const long verticalSpacing = 100000; // 间距 100,000 EMU ≈ 0.11 英寸

            //string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ±2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50Ω\nChip Size: 1.766 x 2.0 x 0.05mm";   
            //AddTextBox(slide, features, 914400, currentY);
            //currentY += textBoxHeight + verticalSpacing;

            //AddTextBox(slide, "第二个文本框标题\n正文行A\n正文行B", 914400, currentY);
            //currentY += textBoxHeight + verticalSpacing;

            //AddTextBox(slide, "第三个文本框\n内容...", 914400, currentY);

            // 插入文本框
            //AddTextBox(slide, "标题行（加粗14号）\n普通行1\n普通行2", 914400, 1214000);
            //AddTextBox(slide, "第二个文本框标题\n正文行A\n正文行B", 914400, 2_000_000);

            // 可选：插入图片
            // if (File.Exists(imagePath))
            //     AddImage(slidePart, imagePath, 5_000_000, 1_000_000, 2_000_000, 2_000_000);



            var currentY = 1214000L; // 初始 Y 位置
            const long verticalSpacing = 500000; // 间距 100,000 EMU

            string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ±2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50Ω\nChip Size: 1.766 x 2.0 x 0.05mm";

            // 添加文本框并获取实际高度
            long height1 = AddTextBox(slide, features, 914400, currentY);
            currentY += height1 + verticalSpacing + 100000;

            string typicalApp = "Test Instrumentation\nMicrowave Radio & VSAT\nMilitary & Space\nTelecom Infrastructure\nFiber Optics";
            long height2 = AddTextBox(slide, typicalApp, 914400, currentY);
            currentY += height2 + verticalSpacing;

            string info2 = "Electrical Specifications";
            long height3 = AddTextBox(slide, info2, 914400, currentY);
            currentY += height3 + 100000;

            string info3 = "TA = +25\u2103, VD = +4V , VG=-0.4V , IDD = 119mA Typical";
            long height4 = AddTextBox(slide, info3, 914400, currentY);
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

            // 计算位置


            // 添加表格
            //AddStyledTable(slide, tableData, 914400, currentY);
            AddTable(slide, tableData, 914400, currentY, 6000000, 3800000);
            currentY += 2000000 + verticalSpacing; // 表格高度 + 间距


            // 添加新幻灯片
            var newSlidePart = AddNewSlideFromLayout(presentationPart);
            //AddNewSlideFromLayout

            var newSlide = newSlidePart.Slide;
            // 示例：在新幻灯片上添加表格
            int originX = 914400;

            int originY = 1314000;
            currentY = originY;
            string info = "Measurement Plots: S-parameters\n TA = +25\u2103";// \u2103 是摄氏度的符号
            long height = AddTextBoxCenter(newSlide, info, originX, originY);
            currentY += height + 50000;

            var offsetX = 914400 + 2_500_000 + 700_000;
            string pic1 = @"F:\PROJECT\ChipManualGeneration\exe\常温\S11.png";
            AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);

            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\常温\S12.png";
            AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);

            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\常温\S21.png";
            currentY += 2_000_000 + 300_000;
            AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);
            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\常温\S22.png";
            AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);

            currentY += 2_000_000 + 250_000;
            info = "Measurement Plots: S-parameters\nVD=4.0V,VG=-0.5V";
            height = AddTextBoxCenter(newSlide, info, originX, currentY);
            currentY += height + 50000;
            AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);
            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\常温\S22.png";
            AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);


            slide.Save();
        }



        //在前面增加了一个圆形项目符号
        public static long AddTextBox(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = i == 0;

                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 对于第二行及以后的行，先添加项目符号
                if (i >= 1)
                {
                    // 创建项目符号的 Run
                    var bulletRunProps = new A.RunProperties
                    {
                        FontSize = 1100,
                    };
                    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                    // 使用Unicode圆点字符，并确保编码正确
                    //var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // Unicode圆点字符后跟一个空格

                    //paragraph.Append(bulletRun);
                }

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    FontSize = isTitle ? 1400 : 1100,
                    Bold = isTitle,
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //string tmp = "?  " + lines[i];

                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                //var textRun = new A.Run(textRunProps, new A.Text(tmp));

                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        public static long AddTextBox2(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = false;

                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 对于第二行及以后的行，先添加项目符号
                //if (i >= 1)
                //{
                //    // 创建项目符号的 Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // 使用Unicode圆点字符，并确保编码正确
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // Unicode圆点字符后跟一个空格

                //    paragraph.Append(bulletRun);
                //}

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    //FontSize = isTitle ? 1400 : 1100,
                    FontSize = 1200,
                    Bold = true,
                    Underline = (i==0? A.TextUnderlineValues.Single:A.TextUnderlineValues.None),
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        public static long AddTextBox8(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = false;

                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 对于第二行及以后的行，先添加项目符号
                //if (i >= 1)
                //{
                //    // 创建项目符号的 Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // 使用Unicode圆点字符，并确保编码正确
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // Unicode圆点字符后跟一个空格

                //    paragraph.Append(bulletRun);
                //}

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    //FontSize = isTitle ? 1400 : 1100,
                    FontSize = 1200,
                    //Bold = true,
                    //Underline = (i == 0 ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None),
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }


        public static long AddTextBox3(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = true;

                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 对于第二行及以后的行，先添加项目符号
                //if (i >= 1)
                //{
                //    // 创建项目符号的 Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // 使用Unicode圆点字符，并确保编码正确
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // Unicode圆点字符后跟一个空格

                //    paragraph.Append(bulletRun);
                //}

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    //FontSize = isTitle ? 1400 : 1100,
                    FontSize = 1200,
                    Bold = isTitle,
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        public static long AddTextBox4(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = i == 0;

                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 对于第二行及以后的行，先添加项目符号
                if (i >= 1)
                {
                    // 创建项目符号的 Run
                    var bulletRunProps = new A.RunProperties
                    {
                        FontSize = 1100,
                    };
                    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                    // 使用Unicode圆点字符，并确保编码正确
                    //var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // Unicode圆点字符后跟一个空格

                    //paragraph.Append(bulletRun);
                }

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    FontSize = 700,
                    Bold = isTitle ? true : false,
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        // 和 AddTextBox2 功能相同，只是字体大小不同
        public static long AddTextBox5(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = false;

                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 对于第二行及以后的行，先添加项目符号
                //if (i >= 1)
                //{
                //    // 创建项目符号的 Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // 使用Unicode圆点字符，并确保编码正确
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // Unicode圆点字符后跟一个空格

                //    paragraph.Append(bulletRun);
                //}

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    //FontSize = isTitle ? 1400 : 1100,
                    FontSize = 1100,
                    Bold = (i==0 ||  i==6 || i== 8 || i==11),
                    
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }


        private static A.RunProperties CreateRunPropertiesFromFontConfig(FontConfig font, bool isTitle = false)
        {
            if (font == null)
                font = new FontConfig(); // 使用默认值

            var props = new A.RunProperties();

            // 字体大小：Open XML 使用 100 * 磅（例如 11pt = 1100）
            int fontSize = font.Size > 0 ? font.Size : (isTitle ? 14 : 11);
            props.FontSize = fontSize * 100;

            // 粗体 / 斜体
            props.Bold = font.IsBold;
            props.Italic = font.IsItalic;

            // 颜色（假设是十六进制如 "#FF0000" 或 "FF0000"）
            if (!string.IsNullOrEmpty(font.Color))
            {
                string colorHex = font.Color.TrimStart('#');
                if (colorHex.Length == 6)
                {
                    //props.SolidFill = new A.SolidFill(new A.RgbColorModelHex(colorHex));
                }
            }

            // 字体族
            string typeface = !string.IsNullOrEmpty(font.Typeface) ? font.Typeface : "Calibri";
            props.Append(new A.LatinFont { Typeface = typeface });

            // 下划线
            if (font.Underline.HasValue)
            {
                props.Underline = font.Underline.Value;
            }

            return props;
        }
        public static long AddTextBoxCenter(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    FontSize = 1400,
                    Bold = true,
                    Underline = A.TextUnderlineValues.Single
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });

                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? 辅助方法：必须和 AddTextBoxCenter 在同一个类中！


        /// <summary>
        /// 和前面一个相比 这个可以改变字体大小
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="text"></param>
        /// <param name="offsetX"></param>
        /// <param name="offsetY"></param>
        /// <param name="fontszie"></param>
        /// <returns></returns>
        public static long AddTextBoxCenter(P.Slide slide, string text, long offsetX, long offsetY, int fontszie)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    FontSize = fontszie,
                    Bold = true,
                    Underline = A.TextUnderlineValues.Single
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });

                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? 辅助方法：必须和 AddTextBoxCenter 在同一个类中！

        /// <summary>
        /// 水平居中显示的文本框
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="text"></param>
        /// <param name="offsetX"></param>
        /// <param name="offsetY"></param>
        /// <returns></returns>
        public static long AddTextBoxCenterWH(P.Slide slide, string text, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = width, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    FontSize = 1200,
                    Bold = true,
                    Underline = A.TextUnderlineValues.Single
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });

                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? 辅助方法：必须和 AddTextBoxCenter 在同一个类中！


        public static long AddTextBoxCenter2(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = 6000000, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    FontSize = 1400,
                    Bold = i == 0 ? true : false,

                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });

                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? 辅助方法：必须和 AddTextBoxCenter 在同一个类中！


        public static long AddTextBoxUnderline(P.Slide slide, string text, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // 视觉属性（位置、大小）
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // 视觉属性（位置、大小）
            var spPr = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = width, Cy = textBoxHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            //var spPr = new P.ShapeProperties(
            //    new A.Transform2D(
            //        new A.Offset { X = offsetX, Y = offsetY },
            //        new A.Extents { Cx = 4000000, Cy = 800000 }
            //    ),
            //    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            //);

            // 创建 P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // 添加段落
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // 创建段落属性
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // 为第二行及以后的行设置缩进和项目符号
                if (i >= 1)
                {
                    // 设置缩进
                    //paragraphProperties.LeftMargin = 360000;  // 整个段落的左缩进
                    paragraphProperties.Indent = -180000;     // 首行缩进（负值表示悬挂缩进，使项目符号突出）
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // 添加文本内容的 Run
                var textRunProps = new A.RunProperties
                {
                    FontSize = 1400,
                    Bold = true,
                    Underline = A.TextUnderlineValues.Single
                };
                textRunProps.Append(new A.LatinFont { Typeface = "Calibri" });

                var textRun = new A.Run(textRunProps, new A.Text(lines[i]));
                paragraph.Append(textRun);

                textBody.Append(paragraph);
            }

            // 创建 Shape 并添加
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? 辅助方法：必须和 AddTextBoxCenter 在同一个类中！

        private static void AddRun(A.Paragraph paragraph, string text, int fontSize, bool bold, int baseline)
        {
            var runProps = new A.RunProperties
            {
                FontSize = fontSize,
                Bold = bold,
                Baseline = baseline
            };
            runProps.Append(new A.LatinFont { Typeface = "Calibri" });
            var textElement = new A.Text(text) { };
            paragraph.Append(new A.Run(runProps, textElement));
        }

        // 根据行数计算文本框高度
        private static long CalculateTextBoxHeight(string[] lines)
        {
            const long baseHeight = 200000; // 基础高度
            const long lineHeight = 150000; // 每行增加的高度

            return baseHeight + (lines.Length * lineHeight);
        }
        public static void AddImage(SlidePart slidePart, string imagePath, long offsetX, long offsetY, long width, long height)
        {
            var imagePart = slidePart.AddImagePart(ImagePartType.Jpeg); // 或 Png
            var stream = File.OpenRead(imagePath);
            imagePart.FeedData(stream);

            string relId = slidePart.GetIdOfPart(imagePart);

            var picture = new P.Picture(
                new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties { Id = 100U, Name = System.IO.Path.GetFileName(imagePath) },
                    new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = BooleanValue.FromBoolean(true) }),
                    new P.ApplicationNonVisualDrawingProperties()
                ),
                new P.BlipFill(
                    new A.Blip { Embed = relId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = offsetX, Y = offsetY },
                        new A.Extents { Cx = width, Cy = height }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            slidePart.Slide.CommonSlideData.ShapeTree.Append(picture);
        }


        private static void AddStyledTable(P.Slide slide, string[,] data, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 生成唯一 ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            long tableWidth = 6000000;
            long tableHeight = 2000000;

            // 非视觉属性
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new S.Table();
            table.Append(new DocumentFormat.OpenXml.Drawing.TableProperties());

            // 表格网格
            var tableGrid = new DocumentFormat.OpenXml.Drawing.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                tableGrid.Append(new A.GridColumn { Width = tableWidth / cols });
            }
            table.Append(tableGrid);

            // 添加行和单元格
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new DocumentFormat.OpenXml.Drawing.TableRow { Height = tableHeight / rows };

                for (int col = 0; col < cols; col++)
                {
                    var tableCell = new DocumentFormat.OpenXml.Drawing.TableCell();

                    // 设置单元格样式
                    var cellProperties = new DocumentFormat.OpenXml.Drawing.TableCellProperties();

                    // 表头行背景色
                    if (row == 0)
                    {
                        cellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "D3D3D3" })); // 浅灰色
                    }

                    // 边框
                    cellProperties.Append(new A.TableCellBorders(
                        //new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }) { Width = 1 }),
                        //new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }) { Width = 1 }),
                        //new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }) { Width = 1 }),
                        //new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }) { Width = 1 })
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }))
                    ));

                    tableCell.Append(cellProperties);

                    // 文本内容
                    var textBody = new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.ParagraphProperties
                            {
                                Alignment = A.TextAlignmentTypeValues.Center
                            },
                            new A.Run(
                                new A.RunProperties
                                {
                                    FontSize = 1100,
                                    Bold = (row == 0),
                                    //LatinFont = new A.LatinFont { Typeface = "Calibri" }
                                },
                                new A.Text(data[row, col] ?? "")
                            )
                        )
                    );

                    tableCell.Append(textBody);
                    tableRow.Append(tableCell);
                }

                table.Append(tableRow);
            }

            // 创建图形框架
            var graphicFrame = new P.GraphicFrame(
                nvSpPr,
                new P.Transform(
                    new A.Offset { X = offsetX, Y = offsetY },
                    new A.Extents { Cx = tableWidth, Cy = tableHeight }
                ),
                new Graphic(
                    new A.GraphicData(
                        table
                    //"http://schemas.openxmlformats.org/drawingml/2006/table"
                    )
                )
            );

            shapeTree.Append(graphicFrame);
        }


        public static void AddTable(P.Slide slide, string[,] data, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            long firstColWidth = width / 5 * 2; // 第一列宽一半
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // 表格网格
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // 均分
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // 行
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    // ? 空单元格用空格占位
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];

                    var tableCell = new A.TableCell();
                    var tableCellProperties = new A.TableCellProperties();

                    // 边框
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // 表头背景
                    if (row == 0)
                    {
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));
                    }

                    //tableCell.Append(tableCellProperties);

                    // ===== 文本体（关键修复区）=====
                    var textBody = new A.TextBody();
                    textBody.Append(new A.BodyProperties {
                           Anchor = A.TextAnchoringTypeValues.Center
                    }); // 可添加 Anchor = A.TextAnchoringTypeValues.Center
                    textBody.Append(new A.ListStyle());

                    var paragraph = new A.Paragraph();
                    paragraph.Append(new A.ParagraphProperties
                    {
                        Alignment = A.TextAlignmentTypeValues.Center
                    });

                    var run = new A.Run();
                    var runProperties = new A.RunProperties
                    {
                        FontSize = (row == 0) ? 1200 : 1100,
                        Bold = true
                    };

                    if (row == 0)
                    {
                        runProperties.Bold = true;
                        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                    }
                    runProperties.Append(new A.LatinFont { Typeface = "Calibri" });
                    run.Append(runProperties);
                    run.Append(new A.Text(cellText)); // cellText 至少是 " "

                    paragraph.Append(run);
                    // ? 必须添加 EndParagraphRunProperties
                    paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });

                    textBody.Append(paragraph);
                    tableCell.Append(textBody);
                    tableCell.Append(tableCellProperties);
                    // ============================

                    tableRow.Append(tableCell);
                }
                table.Append(tableRow);
            }
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic); // ? 不要包装 Graphic

            shapeTree.Append(graphicFrame);
        }


        public static void AddTable6(P.Slide slide, string[,] data, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            long firstColWidth = width / 5 * 2; // 第一列宽一半
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // 表格网格
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // 均分
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // 行
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== 表格单元格属性 ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // 表头背景色
                    if (row == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? 关键点：让表格单元格垂直居中
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    tableCellProperties.AnchorCenter = true;

                    // ==== 文本内容 ====
                    var runProperties = new A.RunProperties
                    {
                        FontSize = (row == 0 ? 1200 : 1100),
                        //Bold = (row == 0 ? true : false),
                        Bold = true
                    };

                    if (row == 0)
                    {
                        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                    }
                    runProperties.Append(new A.LatinFont { Typeface = "Calibri" });
                    var textRun = new A.Run(runProperties, new A.Text(cellText));
                    var textBody = new A.TextBody(
                        new A.BodyProperties
                        {
                            Anchor = A.TextAnchoringTypeValues.Center, // 垂直居中
                            AnchorCenter = true                        // 水平中心锚点
                        },
                        new A.ListStyle(),
                    new A.Paragraph(
                                        new A.ParagraphProperties
                                        {
                                            Alignment = A.TextAlignmentTypeValues.Center // 水平居中
                                        },
                                        textRun,
                                        new A.EndParagraphRunProperties { Language = "en-US" }
                                    )

                    );

                    // 先加文本体，再加属性！
                    tableCell.Append(textBody);
                    tableCell.Append(tableCellProperties);

                    tableRow.Append(tableCell);
                }
                table.Append(tableRow);
            }
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic); // ? 不要包装 Graphic

            shapeTree.Append(graphicFrame);
        }

     

        /// <summary>
        /// 每个单元格的宽度均分， 也就是一样的
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="data"></param>
        /// <param name="offsetX"></param>
        /// <param name="offsetY"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public static void AddTableAverageWidth(P.Slide slide, string[,] data, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            //long firstColWidth = width / 5 * 2; // 第一列宽一半
            //long remainingWidth = width - firstColWidth;
            long ColWidth = width / cols;

            // 表格网格
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // 均分
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = ColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // 行
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    // ? 空单元格用空格占位
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];

                    var tableCell = new A.TableCell();
                    var tableCellProperties = new A.TableCellProperties();

                    // 边框
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // 表头背景
                    if (row == 0)
                    {
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));
                    }

                    //tableCell.Append(tableCellProperties);

                    // ===== 文本体（关键修复区）=====
                    var textBody = new A.TextBody();
                    textBody.Append(new A.BodyProperties()); // 可添加 Anchor = A.TextAnchoringTypeValues.Center
                    textBody.Append(new A.ListStyle());

                    var paragraph = new A.Paragraph();
                    paragraph.Append(new A.ParagraphProperties
                    {
                        Alignment = A.TextAlignmentTypeValues.Center
                    });

                    var run = new A.Run();
                    var runProperties = new A.RunProperties
                    {
                        FontSize = (row == 0) ? 1200 : 1100,
                        Bold = true
                    };

                    if (row == 0)
                    {
                        runProperties.Bold = true;
                        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                    }
                    runProperties.Append(new A.LatinFont { Typeface = "Calibri" });
                    run.Append(runProperties);
                    run.Append(new A.Text(cellText)); // cellText 至少是 " "

                    paragraph.Append(run);
                    // ? 必须添加 EndParagraphRunProperties
                    paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });

                    textBody.Append(paragraph);
                    tableCell.Append(textBody);
                    tableCell.Append(tableCellProperties);
                    // ============================

                    tableRow.Append(tableCell);
                }
                table.Append(tableRow);
            }
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic); // ? 不要包装 Graphic

            shapeTree.Append(graphicFrame);
        }

        public static void AddTable(P.Slide slide, List<(string name, List<string> value, string unit)> data, long offsetX, long offsetY, long width, long height)
        {
            if (data == null || data.Count == 0)
                return;

            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 1. 计算总列数 (cols) 和中间值列数 (middleCols)
            // 假设数据至少有 name 和 unit 两列
            int middleCols = data.Max(d => d.value?.Count ?? 0);
            int cols = 1 + middleCols + 1; // 1: name, middleCols: value, 1: unit
            int rows = data.Count;

            // 2. 确定下一个可用的 ShapeId
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 3. 创建 GraphicFrame 的非视觉属性
            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            // 4. 计算列宽
            // 假设第一列 (Name) 占 20%，最后一列 (Unit) 占 20%，中间列均分 60%
            long nameColWidth = width / 5; // 20%
            long unitColWidth = width / 5; // 20%
            long remainingWidth = width - nameColWidth - unitColWidth;
            long middleColWidth = (middleCols > 0) ? remainingWidth / middleCols : 0;

            // 5. 创建 TableGrid (定义列宽)
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                long colWidth;
                if (col == 0) // Name 列
                {
                    colWidth = nameColWidth;
                }
                else if (col == cols - 1) // Unit 列 (最后一列)
                {
                    colWidth = unitColWidth;
                }
                else // Value 列
                {
                    colWidth = middleColWidth;
                }
                tableGrid.Append(new A.GridColumn { Width = colWidth });
            }
            table.Append(tableGrid);

            // 6. 遍历行数据并创建 TableRow
            for (int row = 0; row < rows; row++)
            {
                // 高度平均分配
                var tableRow = new A.TableRow { Height = height / rows };
                var rowData = data[row];

                // 确保第一行作为表头（如果需要）
                bool isHeaderRow = (row == 0);

                // 列索引 0：Name
                tableRow.Append(CreateTableCell(rowData.name, isHeaderRow, isFirstCol: true));

                // 列索引 1 到 middleCols：Value
                for (int col = 0; col < middleCols; col++)
                {
                    string cellText = (rowData.value != null && col < rowData.value.Count)
                                      ? rowData.value[col]
                                      : "";
                    tableRow.Append(CreateTableCell(cellText, isHeaderRow));
                }

                // 最后一列：Unit
                tableRow.Append(CreateTableCell(rowData.unit, isHeaderRow, isLastCol: true));

                table.Append(tableRow);
            }

            // 7. 组装 GraphicFrame
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic);

            shapeTree.Append(graphicFrame);
        }
        /// <summary>
        /// 辅助方法：创建并返回一个带默认样式的 A.TableCell
        /// </summary>
        private static A.TableCell CreateTableCell(string text, bool isHeader, bool isFirstCol = false, bool isLastCol = false)
        {
            // 空单元格用空格占位
            string cellText = string.IsNullOrEmpty(text) ? " " : text;

            var tableCell = new A.TableCell();
            var tableCellProperties = new A.TableCellProperties();

            // 边框颜色
            string borderColor = "D0D0D0";
            // 表头背景色
            string headerBackColor = "757171";
            // 表头文本颜色
            string headerTextColor = "FFFFFF";

            // 设置边框
            tableCellProperties.Append(new A.TableCellBorders(
                new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor })),
                new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor })),
                new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor })),
                new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor }))
            ));

            // 表头样式
            if (isHeader)
            {
                tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = headerBackColor }));
            }

            // ===== 文本体 =====
            var textBody = new A.TextBody();
            textBody.Append(new A.BodyProperties());
            textBody.Append(new A.ListStyle());

            var paragraph = new A.Paragraph();
            paragraph.Append(new A.ParagraphProperties
            {
                // 文本居中
                Alignment = A.TextAlignmentTypeValues.Center
            });

            var run = new A.Run();
            var runProperties = new A.RunProperties
            {
                FontSize = isHeader ? 1200 : 1100, // 字号
                Bold = isHeader, // 表头加粗
            };

            if (isHeader)
            {
                // 表头文本颜色
                runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = headerTextColor }));
            }

            runProperties.Append(new A.LatinFont { Typeface = "Calibri" });
            run.Append(runProperties);
            run.Append(new A.Text(cellText));

            paragraph.Append(run);
            paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });

            textBody.Append(paragraph);
            tableCell.Append(textBody);
            tableCell.Append(tableCellProperties);

            return tableCell;
        }

        public static void AddTable2(P.Slide slide, string[,] data, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            long firstColWidth = (long)(width / 7 * 4.5); // 第一列宽一半
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // 表格网格
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // 均分
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // 行
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== 表格单元格属性 ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // 表头背景色
                    if (col == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? 关键点：让表格单元格垂直居中
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    tableCellProperties.AnchorCenter = true;
                    

                    // ==== 文本内容 ====
                    var runProperties = new A.RunProperties
                    {
                        FontSize = 1000,
                        Bold = true
                    };

                    if (col == 0)
                    {
                        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                    }

                    var textRun = new A.Run(runProperties, new A.Text(cellText));
                    var textBody = new A.TextBody(
                        new A.BodyProperties
                        {
                            Anchor = A.TextAnchoringTypeValues.Center, // 垂直居中
                            AnchorCenter = true                        // 水平中心锚点
                        },
                        new A.ListStyle(),
                        new A.Paragraph(
                                    new A.ParagraphProperties
                                    {
                                        Alignment = A.TextAlignmentTypeValues.Left // 水平居中
                                    },
                                    textRun,
                                    new A.EndParagraphRunProperties { Language = "en-US" }
                                )

                    );

                    // 先加文本体，再加属性！
                    tableCell.Append(textBody);
                    tableCell.Append(tableCellProperties);

                    tableRow.Append(tableCell);
                }


                table.Append(tableRow);
            }
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic); // ? 不要包装 Graphic

            shapeTree.Append(graphicFrame);
        }
        public static void AddTable7(P.Slide slide, string[,] data, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
              .Where(nv => nv.Id != null)
              .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
              new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
              new P.NonVisualGraphicFrameDrawingProperties(),
              new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            long firstColWidth = (long)(width / 7 * 4.5); // 第一列宽一半
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // 表格网格
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // 均分
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // 行
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== 表格单元格属性 ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                      new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                      new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                      new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                      new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // 表头背景色
                    if (col == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? 关键点：让表格单元格垂直居中
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    //tableCellProperties.AnchorCenter = true;


                    // ==== 文本内容 ====
                    var runProperties = new A.RunProperties
                    {
                        FontSize = 1000,
                        Bold = true
                    };

                    if (col == 0)
                    {
                        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                    }

                    var textRun = new A.Run(runProperties, new A.Text(cellText));
                    var textBody = new A.TextBody(
                      new A.BodyProperties
                      {
                          Anchor = A.TextAnchoringTypeValues.Center, // 垂直居中
                          LeftInset = 0
                      },
                      new A.ListStyle(),
                      new A.Paragraph(
                            new A.ParagraphProperties
                            {
                                Alignment = A.TextAlignmentTypeValues.Left // 水平居中
                            },
                            textRun,
                            new A.EndParagraphRunProperties { Language = "en-US" }
                          )

                    );

                    // 先加文本体，再加属性！
                    tableCell.Append(textBody);
                    tableCell.Append(tableCellProperties);

                    tableRow.Append(tableCell);
                }


                table.Append(tableRow);
            }
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
              new A.Offset { X = offsetX, Y = offsetY },
              new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic); // ? 不要包装 Graphic

            shapeTree.Append(graphicFrame);
        }
        
        /// <summary>
        /// 调整了居中对齐
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="data"></param>
        /// <param name="offsetX"></param>
        /// <param name="offsetY"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public static void AddTable4(P.Slide slide, string[,] data, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            long lastColWidth = width / 7 * 5;
            long remainingWidth = width - lastColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // 表格网格
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // 均分
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == cols - 1) ? lastColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // 行
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== 表格单元格属性 ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // 表头背景色
                    if (row == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? 关键点：让表格单元格垂直居中
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    //tableCellProperties.AnchorCenter = true;

                    // ==== 文本内容 ====
                    var runProperties = new A.RunProperties
                    {
                        FontSize = 1000,
                        Bold = (row ==0 ? true : false),
                        
                    };

                    if (row == 0)
                    {
                        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                    }
                    runProperties.Append(new A.LatinFont { Typeface = "Calibri" });
                    var textRun = new A.Run(runProperties, new A.Text(cellText));
                    var textBody = new A.TextBody(
                        new A.BodyProperties
                        {
                            Anchor = A.TextAnchoringTypeValues.Center, // 垂直居中
                            AnchorCenter = true                        // 水平中心锚点
                        },
                        new A.ListStyle(), 
                    new A.Paragraph(
                                        new A.ParagraphProperties
                                        {
                                            Alignment = A.TextAlignmentTypeValues.Center // 水平居中
                                        },
                                        textRun,
                                        new A.EndParagraphRunProperties { Language = "en-US" }
                                    )
                    //new A.Paragraph(
                    //    new A.ParagraphProperties
                    //    {
                    //        Alignment = A.TextAlignmentTypeValues.Center // 水平居中
                    //    },
                    //    new A.Run(
                    //        new A.RunProperties
                    //        {
                    //            FontSize = 1000,
                    //            Bold = true,

                    //         },
                    //        new A.Text(cellText)
                    //    ),
                    //    new A.EndParagraphRunProperties { Language = "en-US" }
                    //)
                    );

                    // 先加文本体，再加属性！
                    tableCell.Append(textBody);
                    tableCell.Append(tableCellProperties);

                    tableRow.Append(tableCell);
                }


                table.Append(tableRow);
            }
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic); // ? 不要包装 Graphic

            shapeTree.Append(graphicFrame);
        }


        public static void AddTable3(P.Slide slide, string[,] data, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            long lastColWidth = width / 7 * 5; 
            long remainingWidth = width - lastColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // 表格网格
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // 均分
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == cols - 1) ? lastColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // 行
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    // ? 空单元格用空格占位
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];

                    var tableCell = new A.TableCell();
                    var tableCellProperties = new A.TableCellProperties();

                    // 边框
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // 表头背景
                    if (row == 0)
                    {
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));
                    }

                    //tableCell.Append(tableCellProperties);

                    // ===== 文本体（关键修复区）=====
                    var textBody = new A.TextBody();
                    textBody.Append(new A.BodyProperties()); // 可添加 Anchor = A.TextAnchoringTypeValues.Center
                    textBody.Append(new A.ListStyle());

                    var paragraph = new A.Paragraph();
                    paragraph.Append(new A.ParagraphProperties
                    {
                        Alignment = A.TextAlignmentTypeValues.Center
                    });

                    var run = new A.Run();
                    var runProperties = new A.RunProperties
                    {
                        FontSize = (row == 0) ? 1000 : 1000,
                        Bold = true
                    };

                    if (row == 0)
                    {
                        runProperties.Bold = true;
                        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                    }
                    runProperties.Append(new A.LatinFont { Typeface = "Calibri" });
                    run.Append(runProperties);
                    run.Append(new A.Text(cellText)); // cellText 至少是 " "

                    paragraph.Append(run);
                    // ? 必须添加 EndParagraphRunProperties
                    paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });

                    textBody.Append(paragraph);
                    tableCell.Append(textBody);
                    tableCell.Append(tableCellProperties);
                    // ============================

                    tableRow.Append(tableCell);
                }
                table.Append(tableRow);
            }
            var graphicData = new A.GraphicData
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicData.Append(table);

            var graphic = new A.Graphic();
            graphic.Append(graphicData);

            var graphicFrame = new P.GraphicFrame();
            graphicFrame.Append(nvGraphicFramePr);
            graphicFrame.Append(new P.Transform(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = width, Cy = height }
            ));
            graphicFrame.Append(new P.ShapeProperties());
            graphicFrame.Append(graphic); // ? 不要包装 Graphic

            shapeTree.Append(graphicFrame);
        }

        public static SlidePart AddNewSlide(PresentationPart presentationPart)
        {
            // 1. 创建新的 SlidePart
            var slidePart = presentationPart.AddNewPart<SlidePart>();

            // 2. 生成唯一 slide ID
            uint slideId = 256; // PowerPoint 通常从 256 开始
            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList != null && slideIdList.Elements<SlideId>().Any())
            {
                slideId = slideIdList.Elements<SlideId>().Max(s => s.Id.Value) + 1;
            }

            // 3. 创建 Slide 内容（最简结构）
            var slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()
                        ),
                        new P.GroupShapeProperties()
                    )
                ),
                new DocumentFormat.OpenXml.Presentation.ColorMapOverride(
                    new MasterColorMapping()
                )
            );

            slidePart.Slide = slide;

            // 4. 将 SlidePart 关联到 Presentation
            slideIdList.Append(new SlideId { Id = slideId, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

            return slidePart;
        }
        public static SlidePart AddNewSlideFromLayout(PresentationPart presentationPart)
        {
            // 1. 选择一个版式（例如第一个）
            var slideMasterPart = presentationPart.SlideMasterParts.First();
            var layoutPart = slideMasterPart.SlideLayoutParts.First(); // 你可以根据名称选择特定版式

            // 2. 创建新的 SlidePart
            var slidePart = presentationPart.AddNewPart<SlidePart>();

            // 3. ? 关键：克隆 Layout 的 CommonSlideData 到新 Slide
            var layoutCommonSlideData = layoutPart.SlideLayout.CommonSlideData.CloneNode(true) as P.CommonSlideData;

            var slide = new P.Slide(
                layoutCommonSlideData, // 使用 Layout 的结构（包含占位符）
                new P.ColorMapOverride(new MasterColorMapping())
            );

            slidePart.Slide = slide;

            // 4. 关联 LayoutPart（可选但推荐）
            slidePart.AddPart(layoutPart);

            // 5. 添加到 SlideIdList
            uint slideId = 256;
            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList?.Elements<P.SlideId>().Any() == true)
            {
                slideId = slideIdList.Elements<P.SlideId>().Max(s => s.Id.Value) + 1;
            }

            slideIdList.Append(new P.SlideId
            {
                Id = slideId,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

            return slidePart;
        }

        public static void ReplaceTextSlideMasterInPresentation(PresentationDocument doc, string placeholder, string newText)
        {
            foreach (var slideMasterPart in doc.PresentationPart.SlideMasterParts)
            {
                ReplaceTextInPart(slideMasterPart, placeholder, newText);

                // 母版里还有布局 (Slide Layout)
                foreach (var layoutPart in slideMasterPart.SlideLayoutParts)
                {
                    ReplaceTextInPart(layoutPart, placeholder, newText);
                }
            }
        }



        public static void ReplaceAllSlideMasterTextInPresentation(PresentationDocument doc, SliderMasterModel sliderMasterModel)
        {
            PptModifier.ReplaceTextSlideMasterInPresentation(doc, "{Top PN}", sliderMasterModel.TopPN);
            PptModifier.ReplaceTextSlideMasterInPresentation(doc, "{Version}", sliderMasterModel.Version);
            PptModifier.ReplaceTextSlideMasterInPresentation(doc, "{Product Name}", sliderMasterModel.ProductName);
            PptModifier.ReplaceTextSlideMasterInPresentation(doc, "{Frequency Range}", sliderMasterModel.FrequencyRange);
            PptModifier.ReplaceTextSlideMasterInPresentation(doc, "{R PN}", sliderMasterModel.RPN);
            PptModifier.ReplaceTextSlideMasterInPresentation(doc, "{right bar info}", sliderMasterModel.RightBarInfo);
        }

        static void ReplaceTextInPart(OpenXmlPart part, string placeholder, string newText)
        {
            var texts = part.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
            foreach (var text in texts)
            {
                if (text.Text.Contains(placeholder))
                {
                    text.Text = text.Text.Replace(placeholder, newText);
                }
            }
        }
    }

}