using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChipManualGenerationSogt.models
{
    public class TableModel
    {
        public List<List<string>> data;// 用一个二维数组来存储表格数据，外层为行，内层为列
        public List<string> headers;// 表格的头部, 可能没得这个
    }
}
