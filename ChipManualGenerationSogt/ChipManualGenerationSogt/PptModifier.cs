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
                throw new FileNotFoundException("δ�ҵ�PPT�ļ�", filePath);

            var presentationDoc = PresentationDocument.Open(filePath, isEditable: true);
            var presentationPart = presentationDoc.PresentationPart;

            // ��ȡ��һ�� slide part��ͨ����ϵ��
            var slideId = presentationPart.Presentation.SlideIdList?.GetFirstChild<P.SlideId>();
            //var slideId = presentationPart.Presentation.SlideIdList?.FirstOrDefault();
            if (slideId == null)
                throw new InvalidOperationException("PPT��û�лõ�Ƭ��");


            // ? �ؼ���v3.3.0 ���� GetPartById ��ȡ SlidePart�������� OpenXmlPart������תΪ Slide��
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
            var slide = slidePart.Slide;

            //var currentY = 1214000L; // ��ʼ Y λ�ã�1 Ӣ�磩
            //const long textBoxHeight = 800000; // �ı���߶ȣ�EMU��
            //const long verticalSpacing = 100000; // ��� 100,000 EMU �� 0.11 Ӣ��

            //string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ��2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50��\nChip Size: 1.766 x 2.0 x 0.05mm";   
            //AddTextBox(slide, features, 914400, currentY);
            //currentY += textBoxHeight + verticalSpacing;

            //AddTextBox(slide, "�ڶ����ı������\n������A\n������B", 914400, currentY);
            //currentY += textBoxHeight + verticalSpacing;

            //AddTextBox(slide, "�������ı���\n����...", 914400, currentY);

            // �����ı���
            //AddTextBox(slide, "�����У��Ӵ�14�ţ�\n��ͨ��1\n��ͨ��2", 914400, 1214000);
            //AddTextBox(slide, "�ڶ����ı������\n������A\n������B", 914400, 2_000_000);

            // ��ѡ������ͼƬ
            // if (File.Exists(imagePath))
            //     AddImage(slidePart, imagePath, 5_000_000, 1_000_000, 2_000_000, 2_000_000);



            var currentY = 1214000L; // ��ʼ Y λ��
            const long verticalSpacing = 500000; // ��� 100,000 EMU

            string features = "Features\nFrequency: 45-90GHz\nSmall Signal Gain: 15dB Typical\nGain Flatness: ��2.5dB Typical\nNoise Figure:4.5dB Typical\n P1dB: 12dBm Typical\n Power Supply:VD=+4V@119mA ,VG=-0.4V\n Input /Output: 50��\nChip Size: 1.766 x 2.0 x 0.05mm";

            // ����ı��򲢻�ȡʵ�ʸ߶�
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
            { "Gain Flatness",  "", "��1.0", "", "", "��1.0", "", "dB" },
            { "Noise Figure",  "", "��1.0", "", "", "��1.0", "", "dB" },
            { "P1dB - Output 1dB Compression",  "", "12", "", "", "14", "", "dBm" },
            { "Psat - Saturated Output Power",  "", "12", "", "", "14", "", "dBm" },
            { "OIP3 - Output Third Order Intercept",  "", "12", "", "", "14", "", "dBm" },
            { "Input Return Loss",  "", "12", "", "", "14", "", "dB" },
            { "Output Return Loss",  "", "12", "", "", "14", "", "dB" }
                };

            // ����λ��


            // ��ӱ��
            //AddStyledTable(slide, tableData, 914400, currentY);
            AddTable(slide, tableData, 914400, currentY, 6000000, 3800000);
            currentY += 2000000 + verticalSpacing; // ���߶� + ���


            // ����»õ�Ƭ
            var newSlidePart = AddNewSlideFromLayout(presentationPart);
            //AddNewSlideFromLayout

            var newSlide = newSlidePart.Slide;
            // ʾ�������»õ�Ƭ����ӱ��
            int originX = 914400;

            int originY = 1314000;
            currentY = originY;
            string info = "Measurement Plots: S-parameters\n TA = +25\u2103";// \u2103 �����϶ȵķ���
            long height = AddTextBoxCenter(newSlide, info, originX, originY);
            currentY += height + 50000;

            var offsetX = 914400 + 2_500_000 + 700_000;
            string pic1 = @"F:\PROJECT\ChipManualGeneration\exe\����\S11.png";
            AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);

            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\����\S12.png";
            AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);

            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\����\S21.png";
            currentY += 2_000_000 + 300_000;
            AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);
            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\����\S22.png";
            AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);

            currentY += 2_000_000 + 250_000;
            info = "Measurement Plots: S-parameters\nVD=4.0V,VG=-0.5V";
            height = AddTextBoxCenter(newSlide, info, originX, currentY);
            currentY += height + 50000;
            AddImage(newSlidePart, pic1, originX, currentY, 2_500_000, 2_000_000);
            pic1 = @"F:\PROJECT\ChipManualGeneration\exe\����\S22.png";
            AddImage(newSlidePart, pic1, offsetX, currentY, 2_500_000, 2_000_000);


            slide.Save();
        }



        //��ǰ��������һ��Բ����Ŀ����
        public static long AddTextBox(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = i == 0;

                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ���ڵڶ��м��Ժ���У��������Ŀ����
                if (i >= 1)
                {
                    // ������Ŀ���ŵ� Run
                    var bulletRunProps = new A.RunProperties
                    {
                        FontSize = 1100,
                    };
                    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                    // ʹ��UnicodeԲ���ַ�����ȷ��������ȷ
                    //var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // UnicodeԲ���ַ����һ���ո�

                    //paragraph.Append(bulletRun);
                }

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        public static long AddTextBox2(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = false;

                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ���ڵڶ��м��Ժ���У��������Ŀ����
                //if (i >= 1)
                //{
                //    // ������Ŀ���ŵ� Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // ʹ��UnicodeԲ���ַ�����ȷ��������ȷ
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // UnicodeԲ���ַ����һ���ո�

                //    paragraph.Append(bulletRun);
                //}

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        public static long AddTextBox8(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = false;

                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ���ڵڶ��м��Ժ���У��������Ŀ����
                //if (i >= 1)
                //{
                //    // ������Ŀ���ŵ� Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // ʹ��UnicodeԲ���ַ�����ȷ��������ȷ
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // UnicodeԲ���ַ����һ���ո�

                //    paragraph.Append(bulletRun);
                //}

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }


        public static long AddTextBox3(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = true;

                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ���ڵڶ��м��Ժ���У��������Ŀ����
                //if (i >= 1)
                //{
                //    // ������Ŀ���ŵ� Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // ʹ��UnicodeԲ���ַ�����ȷ��������ȷ
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // UnicodeԲ���ַ����һ���ո�

                //    paragraph.Append(bulletRun);
                //}

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        public static long AddTextBox4(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = i == 0;

                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ���ڵڶ��м��Ժ���У��������Ŀ����
                if (i >= 1)
                {
                    // ������Ŀ���ŵ� Run
                    var bulletRunProps = new A.RunProperties
                    {
                        FontSize = 1100,
                    };
                    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                    // ʹ��UnicodeԲ���ַ�����ȷ��������ȷ
                    //var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // UnicodeԲ���ַ����һ���ո�

                    //paragraph.Append(bulletRun);
                }

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }

        // �� AddTextBox2 ������ͬ��ֻ�������С��ͬ
        public static long AddTextBox5(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool isTitle = false;

                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ���ڵڶ��м��Ժ���У��������Ŀ����
                //if (i >= 1)
                //{
                //    // ������Ŀ���ŵ� Run
                //    var bulletRunProps = new A.RunProperties
                //    {
                //        FontSize = 1100,
                //    };
                //    bulletRunProps.Append(new A.LatinFont { Typeface = "Calibri" });
                //    // ʹ��UnicodeԲ���ַ�����ȷ��������ȷ
                //    var bulletRun = new A.Run(bulletRunProps, new A.Text("\u2022 ")); // UnicodeԲ���ַ����һ���ո�

                //    paragraph.Append(bulletRun);
                //}

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        }


        private static A.RunProperties CreateRunPropertiesFromFontConfig(FontConfig font, bool isTitle = false)
        {
            if (font == null)
                font = new FontConfig(); // ʹ��Ĭ��ֵ

            var props = new A.RunProperties();

            // �����С��Open XML ʹ�� 100 * �������� 11pt = 1100��
            int fontSize = font.Size > 0 ? font.Size : (isTitle ? 14 : 11);
            props.FontSize = fontSize * 100;

            // ���� / б��
            props.Bold = font.IsBold;
            props.Italic = font.IsItalic;

            // ��ɫ��������ʮ�������� "#FF0000" �� "FF0000"��
            if (!string.IsNullOrEmpty(font.Color))
            {
                string colorHex = font.Color.TrimStart('#');
                if (colorHex.Length == 6)
                {
                    //props.SolidFill = new A.SolidFill(new A.RgbColorModelHex(colorHex));
                }
            }

            // ������
            string typeface = !string.IsNullOrEmpty(font.Typeface) ? font.Typeface : "Calibri";
            props.Append(new A.LatinFont { Typeface = typeface });

            // �»���
            if (font.Underline.HasValue)
            {
                props.Underline = font.Underline.Value;
            }

            return props;
        }
        public static long AddTextBoxCenter(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? ��������������� AddTextBoxCenter ��ͬһ�����У�


        /// <summary>
        /// ��ǰ��һ����� ������Ըı������С
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

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? ��������������� AddTextBoxCenter ��ͬһ�����У�

        /// <summary>
        /// ˮƽ������ʾ���ı���
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="text"></param>
        /// <param name="offsetX"></param>
        /// <param name="offsetY"></param>
        /// <returns></returns>
        public static long AddTextBoxCenterWH(P.Slide slide, string text, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? ��������������� AddTextBoxCenter ��ͬһ�����У�


        public static long AddTextBoxCenter2(P.Slide slide, string text, long offsetX, long offsetY)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Center
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? ��������������� AddTextBoxCenter ��ͬһ�����У�


        public static long AddTextBoxUnderline(P.Slide slide, string text, long offsetX, long offsetY, long width, long height)
        {
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // ����Ψһ ID
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            // �Ӿ����ԣ�λ�á���С��
            var lines = text.Split('\n');
            long textBoxHeight = CalculateTextBoxHeight(lines);

            // �Ӿ����ԣ�λ�á���С��
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

            // ���� P.TextBody
            var textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );

            // ��Ӷ���
            //var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {


                // ������������
                var paragraphProperties = new A.ParagraphProperties
                {
                    Alignment = A.TextAlignmentTypeValues.Left
                };

                // Ϊ�ڶ��м��Ժ����������������Ŀ����
                if (i >= 1)
                {
                    // ��������
                    //paragraphProperties.LeftMargin = 360000;  // ���������������
                    paragraphProperties.Indent = -180000;     // ������������ֵ��ʾ����������ʹ��Ŀ����ͻ����
                }

                var paragraph = new A.Paragraph(paragraphProperties);

                // ����ı����ݵ� Run
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

            // ���� Shape �����
            var shape = new P.Shape(nvSpPr, spPr, textBody);
            shapeTree.Append(shape);
            return textBoxHeight;
        } // ? ��������������� AddTextBoxCenter ��ͬһ�����У�

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

        // �������������ı���߶�
        private static long CalculateTextBoxHeight(string[] lines)
        {
            const long baseHeight = 200000; // �����߶�
            const long lineHeight = 150000; // ÿ�����ӵĸ߶�

            return baseHeight + (lines.Length * lineHeight);
        }
        public static void AddImage(SlidePart slidePart, string imagePath, long offsetX, long offsetY, long width, long height)
        {
            var imagePart = slidePart.AddImagePart(ImagePartType.Jpeg); // �� Png
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

            // ����Ψһ ID
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

            // ���Ӿ�����
            var nvSpPr = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new S.Table();
            table.Append(new DocumentFormat.OpenXml.Drawing.TableProperties());

            // �������
            var tableGrid = new DocumentFormat.OpenXml.Drawing.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                tableGrid.Append(new A.GridColumn { Width = tableWidth / cols });
            }
            table.Append(tableGrid);

            // ����к͵�Ԫ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new DocumentFormat.OpenXml.Drawing.TableRow { Height = tableHeight / rows };

                for (int col = 0; col < cols; col++)
                {
                    var tableCell = new DocumentFormat.OpenXml.Drawing.TableCell();

                    // ���õ�Ԫ����ʽ
                    var cellProperties = new DocumentFormat.OpenXml.Drawing.TableCellProperties();

                    // ��ͷ�б���ɫ
                    if (row == 0)
                    {
                        cellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "D3D3D3" })); // ǳ��ɫ
                    }

                    // �߿�
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

                    // �ı�����
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

            // ����ͼ�ο��
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

            long firstColWidth = width / 5 * 2; // ��һ�п�һ��
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // �������
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // ����
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    // ? �յ�Ԫ���ÿո�ռλ
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];

                    var tableCell = new A.TableCell();
                    var tableCellProperties = new A.TableCellProperties();

                    // �߿�
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // ��ͷ����
                    if (row == 0)
                    {
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));
                    }

                    //tableCell.Append(tableCellProperties);

                    // ===== �ı��壨�ؼ��޸�����=====
                    var textBody = new A.TextBody();
                    textBody.Append(new A.BodyProperties {
                           Anchor = A.TextAnchoringTypeValues.Center
                    }); // ����� Anchor = A.TextAnchoringTypeValues.Center
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
                    run.Append(new A.Text(cellText)); // cellText ������ " "

                    paragraph.Append(run);
                    // ? ������� EndParagraphRunProperties
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
            graphicFrame.Append(graphic); // ? ��Ҫ��װ Graphic

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

            long firstColWidth = width / 5 * 2; // ��һ�п�һ��
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // �������
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // ����
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== ���Ԫ������ ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // ��ͷ����ɫ
                    if (row == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? �ؼ��㣺�ñ��Ԫ��ֱ����
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    tableCellProperties.AnchorCenter = true;

                    // ==== �ı����� ====
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
                            Anchor = A.TextAnchoringTypeValues.Center, // ��ֱ����
                            AnchorCenter = true                        // ˮƽ����ê��
                        },
                        new A.ListStyle(),
                    new A.Paragraph(
                                        new A.ParagraphProperties
                                        {
                                            Alignment = A.TextAlignmentTypeValues.Center // ˮƽ����
                                        },
                                        textRun,
                                        new A.EndParagraphRunProperties { Language = "en-US" }
                                    )

                    );

                    // �ȼ��ı��壬�ټ����ԣ�
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
            graphicFrame.Append(graphic); // ? ��Ҫ��װ Graphic

            shapeTree.Append(graphicFrame);
        }

     

        /// <summary>
        /// ÿ����Ԫ��Ŀ�Ⱦ��֣� Ҳ����һ����
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

            //long firstColWidth = width / 5 * 2; // ��һ�п�һ��
            //long remainingWidth = width - firstColWidth;
            long ColWidth = width / cols;

            // �������
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // ����
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = ColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    // ? �յ�Ԫ���ÿո�ռλ
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];

                    var tableCell = new A.TableCell();
                    var tableCellProperties = new A.TableCellProperties();

                    // �߿�
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // ��ͷ����
                    if (row == 0)
                    {
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));
                    }

                    //tableCell.Append(tableCellProperties);

                    // ===== �ı��壨�ؼ��޸�����=====
                    var textBody = new A.TextBody();
                    textBody.Append(new A.BodyProperties()); // ����� Anchor = A.TextAnchoringTypeValues.Center
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
                    run.Append(new A.Text(cellText)); // cellText ������ " "

                    paragraph.Append(run);
                    // ? ������� EndParagraphRunProperties
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
            graphicFrame.Append(graphic); // ? ��Ҫ��װ Graphic

            shapeTree.Append(graphicFrame);
        }

        public static void AddTable(P.Slide slide, List<(string name, List<string> value, string unit)> data, long offsetX, long offsetY, long width, long height)
        {
            if (data == null || data.Count == 0)
                return;

            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 1. ���������� (cols) ���м�ֵ���� (middleCols)
            // �������������� name �� unit ����
            int middleCols = data.Max(d => d.value?.Count ?? 0);
            int cols = 1 + middleCols + 1; // 1: name, middleCols: value, 1: unit
            int rows = data.Count;

            // 2. ȷ����һ�����õ� ShapeId
            uint shapeId = 1;
            var existingIds = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Where(nv => nv.Id != null)
                .Select(nv => nv.Id.Value);
            if (existingIds.Any())
                shapeId = (uint)(existingIds.Max() + 1);

            // 3. ���� GraphicFrame �ķ��Ӿ�����
            var nvGraphicFramePr = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Table {shapeId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );

            var table = new A.Table();

            // 4. �����п�
            // �����һ�� (Name) ռ 20%�����һ�� (Unit) ռ 20%���м��о��� 60%
            long nameColWidth = width / 5; // 20%
            long unitColWidth = width / 5; // 20%
            long remainingWidth = width - nameColWidth - unitColWidth;
            long middleColWidth = (middleCols > 0) ? remainingWidth / middleCols : 0;

            // 5. ���� TableGrid (�����п�)
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                long colWidth;
                if (col == 0) // Name ��
                {
                    colWidth = nameColWidth;
                }
                else if (col == cols - 1) // Unit �� (���һ��)
                {
                    colWidth = unitColWidth;
                }
                else // Value ��
                {
                    colWidth = middleColWidth;
                }
                tableGrid.Append(new A.GridColumn { Width = colWidth });
            }
            table.Append(tableGrid);

            // 6. ���������ݲ����� TableRow
            for (int row = 0; row < rows; row++)
            {
                // �߶�ƽ������
                var tableRow = new A.TableRow { Height = height / rows };
                var rowData = data[row];

                // ȷ����һ����Ϊ��ͷ�������Ҫ��
                bool isHeaderRow = (row == 0);

                // ������ 0��Name
                tableRow.Append(CreateTableCell(rowData.name, isHeaderRow, isFirstCol: true));

                // ������ 1 �� middleCols��Value
                for (int col = 0; col < middleCols; col++)
                {
                    string cellText = (rowData.value != null && col < rowData.value.Count)
                                      ? rowData.value[col]
                                      : "";
                    tableRow.Append(CreateTableCell(cellText, isHeaderRow));
                }

                // ���һ�У�Unit
                tableRow.Append(CreateTableCell(rowData.unit, isHeaderRow, isLastCol: true));

                table.Append(tableRow);
            }

            // 7. ��װ GraphicFrame
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
        /// ��������������������һ����Ĭ����ʽ�� A.TableCell
        /// </summary>
        private static A.TableCell CreateTableCell(string text, bool isHeader, bool isFirstCol = false, bool isLastCol = false)
        {
            // �յ�Ԫ���ÿո�ռλ
            string cellText = string.IsNullOrEmpty(text) ? " " : text;

            var tableCell = new A.TableCell();
            var tableCellProperties = new A.TableCellProperties();

            // �߿���ɫ
            string borderColor = "D0D0D0";
            // ��ͷ����ɫ
            string headerBackColor = "757171";
            // ��ͷ�ı���ɫ
            string headerTextColor = "FFFFFF";

            // ���ñ߿�
            tableCellProperties.Append(new A.TableCellBorders(
                new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor })),
                new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor })),
                new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor })),
                new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = borderColor }))
            ));

            // ��ͷ��ʽ
            if (isHeader)
            {
                tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = headerBackColor }));
            }

            // ===== �ı��� =====
            var textBody = new A.TextBody();
            textBody.Append(new A.BodyProperties());
            textBody.Append(new A.ListStyle());

            var paragraph = new A.Paragraph();
            paragraph.Append(new A.ParagraphProperties
            {
                // �ı�����
                Alignment = A.TextAlignmentTypeValues.Center
            });

            var run = new A.Run();
            var runProperties = new A.RunProperties
            {
                FontSize = isHeader ? 1200 : 1100, // �ֺ�
                Bold = isHeader, // ��ͷ�Ӵ�
            };

            if (isHeader)
            {
                // ��ͷ�ı���ɫ
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

            long firstColWidth = (long)(width / 7 * 4.5); // ��һ�п�һ��
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // �������
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // ����
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== ���Ԫ������ ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // ��ͷ����ɫ
                    if (col == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? �ؼ��㣺�ñ��Ԫ��ֱ����
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    tableCellProperties.AnchorCenter = true;
                    

                    // ==== �ı����� ====
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
                            Anchor = A.TextAnchoringTypeValues.Center, // ��ֱ����
                            AnchorCenter = true                        // ˮƽ����ê��
                        },
                        new A.ListStyle(),
                        new A.Paragraph(
                                    new A.ParagraphProperties
                                    {
                                        Alignment = A.TextAlignmentTypeValues.Left // ˮƽ����
                                    },
                                    textRun,
                                    new A.EndParagraphRunProperties { Language = "en-US" }
                                )

                    );

                    // �ȼ��ı��壬�ټ����ԣ�
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
            graphicFrame.Append(graphic); // ? ��Ҫ��װ Graphic

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

            long firstColWidth = (long)(width / 7 * 4.5); // ��һ�п�һ��
            long remainingWidth = width - firstColWidth;
            long otherColWidth = remainingWidth / (cols - 1);

            // �������
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // ����
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == 0) ? firstColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== ���Ԫ������ ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                      new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                      new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                      new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                      new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // ��ͷ����ɫ
                    if (col == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? �ؼ��㣺�ñ��Ԫ��ֱ����
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    //tableCellProperties.AnchorCenter = true;


                    // ==== �ı����� ====
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
                          Anchor = A.TextAnchoringTypeValues.Center, // ��ֱ����
                          LeftInset = 0
                      },
                      new A.ListStyle(),
                      new A.Paragraph(
                            new A.ParagraphProperties
                            {
                                Alignment = A.TextAlignmentTypeValues.Left // ˮƽ����
                            },
                            textRun,
                            new A.EndParagraphRunProperties { Language = "en-US" }
                          )

                    );

                    // �ȼ��ı��壬�ټ����ԣ�
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
            graphicFrame.Append(graphic); // ? ��Ҫ��װ Graphic

            shapeTree.Append(graphicFrame);
        }
        
        /// <summary>
        /// �����˾��ж���
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

            // �������
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // ����
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == cols - 1) ? lastColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];
                    var tableCell = new A.TableCell();

                    // ==== ���Ԫ������ ====
                    var tableCellProperties = new A.TableCellProperties();
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // ��ͷ����ɫ
                    if (row == 0)
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));

                    // ?? �ؼ��㣺�ñ��Ԫ��ֱ����
                    tableCellProperties.Anchor = A.TextAnchoringTypeValues.Center;
                    //tableCellProperties.AnchorCenter = true;

                    // ==== �ı����� ====
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
                            Anchor = A.TextAnchoringTypeValues.Center, // ��ֱ����
                            AnchorCenter = true                        // ˮƽ����ê��
                        },
                        new A.ListStyle(), 
                    new A.Paragraph(
                                        new A.ParagraphProperties
                                        {
                                            Alignment = A.TextAlignmentTypeValues.Center // ˮƽ����
                                        },
                                        textRun,
                                        new A.EndParagraphRunProperties { Language = "en-US" }
                                    )
                    //new A.Paragraph(
                    //    new A.ParagraphProperties
                    //    {
                    //        Alignment = A.TextAlignmentTypeValues.Center // ˮƽ����
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

                    // �ȼ��ı��壬�ټ����ԣ�
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
            graphicFrame.Append(graphic); // ? ��Ҫ��װ Graphic

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

            // �������
            var tableGrid = new A.TableGrid();
            for (int col = 0; col < cols; col++)
            {
                // ����
                //tableGrid.Append(new A.GridColumn { Width = width / cols });
                long colWidth = (col == cols - 1) ? lastColWidth : otherColWidth;
                tableGrid.Append(new A.GridColumn { Width = colWidth });
                //var tableRow = new A.TableRow { Height = 100000 };
            }
            table.Append(tableGrid);

            // ��
            for (int row = 0; row < rows; row++)
            {
                var tableRow = new A.TableRow { Height = height / rows };
                //var tableRow = new A.TableRow { Height = 500000 };
                for (int col = 0; col < cols; col++)
                {
                    // ? �յ�Ԫ���ÿո�ռλ
                    string cellText = string.IsNullOrEmpty(data[row, col]) ? " " : data[row, col];

                    var tableCell = new A.TableCell();
                    var tableCellProperties = new A.TableCellProperties();

                    // �߿�
                    tableCellProperties.Append(new A.TableCellBorders(
                        new A.LeftBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.RightBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.TopBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" })),
                        new A.BottomBorder(new A.SolidFill(new A.RgbColorModelHex { Val = "D0D0D0" }))
                    ));

                    // ��ͷ����
                    if (row == 0)
                    {
                        tableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "757171" }));
                    }

                    //tableCell.Append(tableCellProperties);

                    // ===== �ı��壨�ؼ��޸�����=====
                    var textBody = new A.TextBody();
                    textBody.Append(new A.BodyProperties()); // ����� Anchor = A.TextAnchoringTypeValues.Center
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
                    run.Append(new A.Text(cellText)); // cellText ������ " "

                    paragraph.Append(run);
                    // ? ������� EndParagraphRunProperties
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
            graphicFrame.Append(graphic); // ? ��Ҫ��װ Graphic

            shapeTree.Append(graphicFrame);
        }

        public static SlidePart AddNewSlide(PresentationPart presentationPart)
        {
            // 1. �����µ� SlidePart
            var slidePart = presentationPart.AddNewPart<SlidePart>();

            // 2. ����Ψһ slide ID
            uint slideId = 256; // PowerPoint ͨ���� 256 ��ʼ
            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList != null && slideIdList.Elements<SlideId>().Any())
            {
                slideId = slideIdList.Elements<SlideId>().Max(s => s.Id.Value) + 1;
            }

            // 3. ���� Slide ���ݣ����ṹ��
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

            // 4. �� SlidePart ������ Presentation
            slideIdList.Append(new SlideId { Id = slideId, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

            return slidePart;
        }
        public static SlidePart AddNewSlideFromLayout(PresentationPart presentationPart)
        {
            // 1. ѡ��һ����ʽ�������һ����
            var slideMasterPart = presentationPart.SlideMasterParts.First();
            var layoutPart = slideMasterPart.SlideLayoutParts.First(); // ����Ը�������ѡ���ض���ʽ

            // 2. �����µ� SlidePart
            var slidePart = presentationPart.AddNewPart<SlidePart>();

            // 3. ? �ؼ�����¡ Layout �� CommonSlideData ���� Slide
            var layoutCommonSlideData = layoutPart.SlideLayout.CommonSlideData.CloneNode(true) as P.CommonSlideData;

            var slide = new P.Slide(
                layoutCommonSlideData, // ʹ�� Layout �Ľṹ������ռλ����
                new P.ColorMapOverride(new MasterColorMapping())
            );

            slidePart.Slide = slide;

            // 4. ���� LayoutPart����ѡ���Ƽ���
            slidePart.AddPart(layoutPart);

            // 5. ��ӵ� SlideIdList
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

                // ĸ���ﻹ�в��� (Slide Layout)
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