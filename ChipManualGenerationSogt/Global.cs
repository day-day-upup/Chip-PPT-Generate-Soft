using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ChipManualGenerationSogt
{
    public static class Global
    {
        public static User User { get; set; }

        public static TaskTableItem TaskModel { get; set; }

        public static OperationModel OperationModel { get; set; }

        public static string  AppBaseUrl { get; set; }  = AppDomain.CurrentDomain.BaseDirectory;

        public static string  FileBasePath { get; set; }  = System.IO.Path.Combine(AppBaseUrl, "resources","files");

        public static string  FtpRootPath { get; set; }  = "Manuals";
    }
}
