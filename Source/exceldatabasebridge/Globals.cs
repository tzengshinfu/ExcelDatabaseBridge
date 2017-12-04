using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = NetOffice.ExcelApi;

namespace ExcelDatabaseBridge {
    public static class Globals {
        public static Excel.Application app;
        public static Excel.Workbook book;
        public static Excel.Worksheet sheet;
    }
}
