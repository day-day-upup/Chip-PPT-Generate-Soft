using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChipManualGenerationSogt
{
    // ���������ڼ�¼ɸѡ������ �������ջᱣ�浽������ȥ
    // ��Ӧ��json���ݸ�ʽ
//    {
//  "PN": "Example_PN_123",
//  "ON": "Order_A456",
//  "StartDateTime": "2025-10-23T10:00:00Z",
//  "StopDateTime": "2025-10-23T11:00:00Z",
//  "VD_VG_Conditon": [
//    "VD=1V",
//    "VG=2V",
//    "Temp=25C"
//  ],
//  "Min": 10.5,
//  "Max": 99.99
//}
public class FileterConditionModel
    {
        public string PN { set; get; }

        public string ON { set; get; }


        public DateTime? StartDateTime { set; get; }


        public DateTime?  StopDateTime { set; get; }


        public List<string> VD_VG_Conditon { set; get; } = new List<string>();

        public List<string> FreqBands { set; get; } = new List<string>();
        public double  Min { set; get; }
        public double  Max { set; get; }

    }
}
