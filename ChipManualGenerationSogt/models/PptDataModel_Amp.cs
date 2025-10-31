using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace ChipManualGenerationSogt.models
{
    // ��������������洢PPT�зŴ��� amplifier���������

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
        //���
        public string[,] ParameterTableData { get; set; }
        public ImageModel FunctionalBlockDiagramImage { get; set; } = new ImageModel();
    }

    //��������ҳ
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
        //���
    }


    public class EndToFront2
    {
        public ImageModel StructImage { get; set; }

        public string Title { get; set; } = "Biasing and Operation";
        public string TurnOn { get; set; }
        public string TurnOff { get; set; }
        //���
    }


    public class LastPage
    {
        public ImageModel Image { get; set; }
        public string Text1 { get; set; }
        public string Text2 { get; set; }
        //���
    }

    public class CurvesImagePageModel 
    {
        // �ڶ�ҳ�������ҳ�� ȫ�Ǵ洢���ߵ�
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
        // ĸ�������Ϣ
       public SliderMasterModel SliderMaster { get; set; }

        // ��һҳ����
       public  FirstPageModel FirstPage { get; set; }


        // ����ҳ
        public CurvesImagePageModel CurvesImagePage { get; set; }



        //�µ�ҳ
        public EndToFront5 EndToFront5Page { get; set; }


        //����ҳ

        public EndToFront4 EndToFront4Page { get; set; }

        //��������ҳ
       
        public EndToFront3 EndToFront3Page { get; set; }

        // �����ڶ�ҳ
       public EndToFront2 EndToFront2Page { get; set; }

        //���һҳ
        public LastPage LastPage { get; set; }


        
    }
   
}
