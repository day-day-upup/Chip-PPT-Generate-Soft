using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace ChipManualGenerationSogt.models
{
    // 这个个类是用来存储PPT中放大器 amplifier的相关数据

    public class SliderMasterModel
    {
        public string TopPN { get; set; } = "MML806";
        public string Version { get; set; } = "V1.0.0";
        public string ProductName { get; set; } = "V1.0.0";
        public string FrequencyRange { get; set; } = "45-90";
        public string RPN { get; set; } = "MML806";
        public string RightBarInfo { get; set; } = "V1.0.0";

    }

    public class FirstPageModel
    {
        public string FeaturesText { get; set; }
        public string TypicalApplicationsText { get; set; } 
        public string ElectricalSpecsTitle { get; set; }
        public string ElectricalSpecsCondition { get; set; }
        //表格
        public string[,] ParameterTableData { get; set; }
        public ImageModel FunctionalBlockDiagramImage { get; set; } = new ImageModel();
    }

    //倒数第五页
    public class EndToFront5
    {
        public string AbsoluteMaximumRatingsTableTitle { get; set; } = "Absolute Maximum Ratings";
        public string[,] AbsoluteMaximumRatingsTable { get; set; }

        public string TypicalSupplyCurrentVgTableTitle { get; set; } = "Typical Supply Current";
        public string[,] TypicalSupplyCurrentVgTable { get; set; }

        public ImageModel WarningImage { get; set; } 

        public string WarningText { get; set; } = "ELECTROSTATIC SENSITIVE DEVICE OBSERVE HANDLING PRECAUTIONS";
    }


    public class EndToFront4
    {
        public ImageModel PinImage { get; set; } 
        public string NoteText { get; set; }
    }


    public class EndToFront3
    {
        public ImageModel StructImage { get; set; }
        public string[,] Description { get; set; }
        public string[,] Description2 { get; set; }
        //表格
    }


    public class EndToFront2
    {
        public ImageModel StructImage { get; set; }

        public string Title { get; set; } = "Biasing and Operation";
        public string TurnOn { get; set; }
        public string TurnOff { get; set; }
        //表格
    }


    public class LastPage
    {
        public ImageModel Image { get; set; }
        public string Text1 { get; set; }
        public string Text2 { get; set; }
        //表格
    }

    public class CurvesImagePageModel 
    {
        // 第二页到后面的页， 全是存储曲线的
        public List<string> CurveTitles { get; set; } = new List<string>();
        public List<string> CurveImagesPath { get; set; } =  new List<string>();

    }

    public class ImageModel
    {
        public string ImagePath { get; set; }
        public string ImageName { get; set; }

        public decimal Width { get; set; } = 2500;
        public decimal Height { get; set; }

        public decimal XPoistion { get; set; }
        public decimal YPoistion { get; set; }

    }
    public class PptDataModel
    {
        // 母版基本信息
       public SliderMasterModel SliderMaster { get; set; }

        // 第一页内容
       public  FirstPageModel FirstPage { get; set; }


        // 曲线页
        public CurvesImagePageModel CurvesImagePage { get; set; }



        //新的页
        public EndToFront5 EndToFront5Page { get; set; }


        //引脚页

        public EndToFront4 EndToFront4Page { get; set; }

        //倒数第三页
       
        public EndToFront3 EndToFront3Page { get; set; }

        // 倒数第二页
       public EndToFront2 EndToFront2Page { get; set; }

        //最后一页
        public LastPage LastPage { get; set; }


        
    }
   
}
