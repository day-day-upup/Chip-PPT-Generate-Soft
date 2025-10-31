using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChipManualGenerationSogt.models
{
    /// <summary>
    /// ר����������Ŵ��������������
    /// һ��һά����
    /// ������ά����-1.�����µ�S11��S12,S22,S21, NF,Psat, Pxdb, OIP3
    ///             -2 ��ͬVG�µ�S11��S12,S22,S21, NF,Psat, Pxdb, OIP3
    /// </summary>

    //public class AmpCurveModel
    //{
    //    public PlotModel[]  StandardSParameters { get; set; } = new PlotModel[4]; // 4 ��s����ͼ�� 0-S11, 1-S12, 2-S21, 3-S22
    //    public List<ThreeTemperaturePlotModel> ThreeTemperaturePlotModel { set; get; }// ���±�


    //    public Temperature25CPlotModel Temperature25CPlotModel { set; get; } // 25��ʱ ��ͬVG�γɵı�
    //}

    //public class Temperature25CPlotModel
    //{
    //    // 25��ʱ�����ߣ� ÿ����ͬ�����¿����ж��VG�� ÿ����VGΪһ�飬 �����ж��飬 ���ܾͻ��γ�����
    //    public List<PlotModel> S11List { set; get; } = new List<PlotModel>();
    //    public List<PlotModel> S12List { set; get; } = new List<PlotModel>();
    //    public List<PlotModel> S21List { set; get; } = new List<PlotModel>();
    //    public List<PlotModel> S22List { set; get; } = new List<PlotModel>();
    //    public List<PlotModel> NFList { set; get; } = new List<PlotModel>();
    //    public List<PlotModel> PsatList { set; get; } = new List<PlotModel>();
    //    public List<PlotModel> PxdbList { set; get; } = new List<PlotModel>();
    //    public List<PlotModel> OIP3List { set; get; } = new List<PlotModel>();


    //    //public List<PlotModel> S11List { set; get; } = new List<PlotModel>();
    //    //public List<PlotModel> S12List { set; get; } = new List<PlotModel>();
    //    //public List<PlotModel> S21List { set; get; } = new List<PlotModel>();
    //    //public List<PlotModel> S22List { set; get; } = new List<PlotModel>();
    //    //public List<PlotModel> NFList { set; get; } = new List<PlotModel>();
    //    //public List<PlotModel> PsatList { set; get; } = new List<PlotModel>();
    //    //public List<PlotModel> PxdbList { set; get; } = new List<PlotModel>();
    //    //public List<PlotModel> OIP3List { set; get; } = new List<PlotModel>();

    //}

    /// <summary>
    /// 25�� ���� �ļ�·��ģ��
    /// </summary>
    public class Temperature25FilePathModel
    {
        // 25��ʱ�����ߣ� ÿ����ͬ�����¿����ж��VG�� ÿ����VGΪһ�飬 �����ж��飬 ���ܾͻ��γ�����
        //public List<PlotModel> S11List { set;get; } = new List<PlotModel>();
        //public List<PlotModel> S12List { set;get; } = new List<PlotModel>();
        //public List<PlotModel> S21List { set;get; } = new List<PlotModel>();
        //public List<PlotModel> S22List { set;get; } = new List<PlotModel>();
        //public List<PlotModel> NFList { set;get; } = new List<PlotModel>();
        //public List<PlotModel> PsatList { set;get; } = new List<PlotModel>();
        //public List<PlotModel> PxdbList { set;get; } = new List<PlotModel>();
        //public List<PlotModel> OIP3List { set;get; } = new List<PlotModel>();

        public List<string> SList { set; get; } = new List<string>();
        public List<string> NFList { set; get; } = new List<string>();
        public List<string> PsatList { set; get; } = new List<string>();
        public List<string> PxdbList { set; get; } = new List<string>();
        public List<string> OIP3List { set; get; } = new List<string>();

    }

    //public class ThreeTemperaturePlotModel
    //{
    //    public PlotModel S11 { set; get; }
    //    public PlotModel S12 { set; get; }

    //    public PlotModel S21 { set; get; }
    //    public PlotModel S22 { set; get; }
            
    //    public PlotModel NF { set; get; }
    //    public PlotModel Psat { set; get; }
    //    public PlotModel Pxdb { set; get; }
    //    public PlotModel OIP3 { set; get; }



    //}


}
