using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = NetOffice.ExcelApi;

namespace ExcelDatabaseBridge {
    public static class ExtensionMethod {
        public static void BeginUpdate(this Excel.Application app) {
            app.ScreenUpdating = false;
            app.DisplayStatusBar = false;
            app.Calculation = Excel.Enums.XlCalculation.xlCalculationManual;
            app.EnableEvents = false;
        }

        public static void EndUpdate(this Excel.Application app) {
            app.ScreenUpdating = true;
            app.DisplayStatusBar = true;
            app.Calculation = Excel.Enums.XlCalculation.xlCalculationAutomatic;
            app.EnableEvents = true;
        }

        public static bool IsValid(this object obj) {
            bool result = true;

            if (obj == null) {
                result = false;
            }
            else if (obj == DBNull.Value) {
                result = false;
            }
            else {
                string text = obj.ToString();
                if (string.IsNullOrWhiteSpace(text) == true) {
                    result = false;
                }
            }

            return result;
        }

        public static void ForEach<T>(this IEnumerable<T> enumeration, Action<T> action) {
            foreach (T item in enumeration) {
                action(item);
            }
        }

        public static string FormatWithArgs(this string str, params object[] args) {
            return string.Format(str, args);
        }
    }
}