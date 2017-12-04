using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = NetOffice.ExcelApi;

namespace ExcelDatabaseBridge {
    public class Main : IExcelAddIn {
        public void AutoOpen() {
            Globals.app = new Excel.Application(null, ExcelDnaUtil.Application);
            Globals.app.WorkbookActivateEvent += WorkbookActivateEvent;
            Globals.app.SheetActivateEvent += WorksheetActivateEvent;
        }

        void WorksheetActivateEvent(NetOffice.COMObject Sh) {
            Globals.sheet = (Excel.Worksheet)Sh;
        }

        void WorkbookActivateEvent(Excel.Workbook Wb) {
            Globals.book = Wb;
            Globals.sheet = (Excel.Worksheet)Wb.ActiveSheet;
        }

        public void AutoClose() {

        }
    }

    [ComVisible(true)]
    public class RibbonController : ExcelRibbon {
        private DataTable fields;
        private string currentHostName;
        private string currentBookName;
        private string currentSheetName;
        private string currentUserName;
        private string currentFilePath;
        private string ribbonHostName;
        private string ribbonBookName;
        private string ribbonSheetName;
        private string ribbtonModifiedDatetime;

        public override string GetCustomUI(string RibbonID) {
            string menu = @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                               <ribbon>
                                   <tabs>
                                       <tab id='tab1' label='{0}'>
                                           <group id='group1' label='{1}'>
                                               <button id='button_ExcelToDatabase' size='large' label='{2}' onAction='button_ExcelToDatabase_Click' getImage='GetImage' />
                                           </group >
                                           <group id='group2' label='{3}'>
                                               <comboBox id='comboBox_HostName' label='{5}' invalidateContentOnDrop='true' getItemLabel='GetItemLabel' getItemCount='GetItemCount' onChange='GetItemText' />
                                               <comboBox id='comboBox_BookName' label='{6}' invalidateContentOnDrop='true' getItemLabel='GetItemLabel' getItemCount='GetItemCount' onChange='GetItemText' />
                                               <comboBox id='comboBox_SheetName' label='{7}' invalidateContentOnDrop='true' getItemLabel='GetItemLabel' getItemCount='GetItemCount' onChange='GetItemText' />
                                               <comboBox id='comboBox_ModifiedDatetime' label='{8}' invalidateContentOnDrop='true' getItemLabel='GetItemLabel' getItemCount='GetItemCount' onChange='GetItemText' />
                                               <button id='button_DatabaseToExcel' size='large' label='{4}' onAction='button_DatabaseToExcel_Click' getImage='GetImage' />
                                           </group >
                                       </tab>
                                   </tabs>
                               </ribbon>
                            </customUI>";
            menu = menu.FormatWithArgs("ExcelDatabaseBridge", "寫入資料庫", "執行", "讀取資料庫", "執行", "主機名稱", "檔案名稱", "工作表名稱", "存檔時間");

            return menu;
        }

        public int GetItemCount(IRibbonControl control) {
            Database.InitialCommand();

            int result = 0;

            fields = GetFields(control.Id);
            result = fields.Rows.Count;

            return result;
        }

        public string GetItemLabel(IRibbonControl control, int index) {
            Database.DisposeCommand();

            return fields.Rows[index][0].ToString();
        }

        public void button_ExcelToDatabase_Click(IRibbonControl control) {
            SetEnvironmentVariable();
            Database.InitialCommand();

            DataTable bulkCopyTable = Database.GetDataTableFromSQL(@"SELECT * FROM ""edb_saved_excel"" WHERE 1 = 0;").Clone();
            DateTime modifiedDatetime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Kind);

            var usedCells = Globals.sheet.UsedRange;
            foreach (var usedCell in usedCells) {
                DataRow bulkCopyRow = bulkCopyTable.NewRow();
                bulkCopyRow["host_name"] = currentHostName;
                bulkCopyRow["file_path"] = currentFilePath;
                bulkCopyRow["book_name"] = currentBookName;
                bulkCopyRow["sheet_name"] = currentSheetName;
                bulkCopyRow["user_account"] = currentUserName;
                bulkCopyRow["modified_datetime"] = modifiedDatetime;
                bulkCopyRow["row"] = usedCell.Row;
                bulkCopyRow["col"] = usedCell.Column;
                bulkCopyRow["value"] = usedCell.Value2;
                bulkCopyRow["format"] = usedCell.NumberFormat.ToString();
                bulkCopyRow["formula"] = usedCell.Formula.ToString();

                bulkCopyTable.Rows.Add(bulkCopyRow);
            }

            Database.InitialTransaction();
            Database.command.ExecuteNonQuery();

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(Database.conn, SqlBulkCopyOptions.KeepIdentity, Database.transaction)) {
                bulkCopy.BatchSize = 1000;
                bulkCopy.BulkCopyTimeout = 60;
                bulkCopy.DestinationTableName = "edb_saved_excel";

                bulkCopy.ColumnMappings.Add("host_name", "host_name");
                bulkCopy.ColumnMappings.Add("file_path", "file_path");
                bulkCopy.ColumnMappings.Add("book_name", "book_name");
                bulkCopy.ColumnMappings.Add("sheet_name", "sheet_name");
                bulkCopy.ColumnMappings.Add("user_account", "user_account");
                bulkCopy.ColumnMappings.Add("modified_datetime", "modified_datetime");
                bulkCopy.ColumnMappings.Add("row", "row");
                bulkCopy.ColumnMappings.Add("col", "col");
                bulkCopy.ColumnMappings.Add("value", "value");
                bulkCopy.ColumnMappings.Add("format", "format");
                bulkCopy.ColumnMappings.Add("formula", "formula");

                bulkCopy.WriteToServer(bulkCopyTable);

                try {
                    Database.transaction.Commit();

                    Database.DisposeCommand();
                    MessageBox.Show("儲存成功");
                }
                catch (SqlException ex) {
                    Database.transaction.Rollback();

                    Database.DisposeCommand();
                    MessageBox.Show("錯誤:" + ":" + Environment.NewLine + ex.Message);
                }
            }
        }

        public void button_DatabaseToExcel_Click(IRibbonControl control) {
            Database.InitialCommand();

            IEnumerable<DataRow> savedExcelRecords = Database.GetDataTableFromSQL(@"SELECT ""row"", ""col"", ""value"", ""format"", ""formula""
                                                                          FROM ""edb_saved_excel""
                                                                          WHERE ""host_name"" = '{0}'
                                                                              AND ""book_name"" = '{1}'
                                                                              AND ""sheet_name"" = '{2}'
                                                                              AND ""modified_datetime"" = '{3}'
                                                                          ORDER BY ""row"", ""col"";".FormatWithArgs(ribbonHostName, ribbonBookName, ribbonSheetName, ribbtonModifiedDatetime)).AsEnumerable();

            Database.DisposeCommand();

            int maxSavedExcelRowCount = savedExcelRecords.Max(e => e.Field<int>("row"));
            int maxSavedExcelColumnCount = savedExcelRecords.Max(e => e.Field<int>("col"));
            object[,] savedExcelValueArray = new object[maxSavedExcelRowCount, maxSavedExcelColumnCount];

            Parallel.ForEach(savedExcelRecords, edbSavedExcelRecord => {
                int savedExcelValueArrayRowIndex = edbSavedExcelRecord.Field<int>("row") - 1;
                int savedExcelValueArrayColumnIndex = edbSavedExcelRecord.Field<int>("col") - 1;

                //如果有設定公式則取公式的值
                savedExcelValueArray[savedExcelValueArrayRowIndex, savedExcelValueArrayColumnIndex] = edbSavedExcelRecord["value"] == edbSavedExcelRecord["formula"] ? edbSavedExcelRecord["value"] : edbSavedExcelRecord["formula"];
             });

            var startPositionCell = Globals.sheet.Range("$A$1");
            var endPositionCell = startPositionCell.Offset(maxSavedExcelRowCount - 1, maxSavedExcelColumnCount - 1);
            var pasteExcelRange = Globals.sheet.Range(startPositionCell, endPositionCell);

            Globals.app.BeginUpdate();
            Globals.sheet.UsedRange.ClearContents();
            pasteExcelRange.Value2 = savedExcelValueArray;
            Globals.app.EndUpdate();
        }

        private void SetEnvironmentVariable() {
            currentHostName = Environment.MachineName;
            currentBookName = Globals.book.Name;
            currentSheetName = Globals.sheet.Name;
            currentUserName = Environment.UserName;
            currentFilePath = Globals.book.Path;
        }

        public void GetItemText(IRibbonControl control, string text) {
            switch (control.Id) {
                case "comboBox_HostName":
                    ribbonHostName = text;

                    break;
                case "comboBox_BookName":
                    ribbonBookName = text;

                    break;
                case "comboBox_SheetName":
                    ribbonSheetName = text;

                    break;
                case "comboBox_ModifiedDatetime":
                    ribbtonModifiedDatetime = text;

                    break;
            }
        }

        public Bitmap GetImage(IRibbonControl control) {
            switch (control.Id) {
                case "button_ExcelToDatabase":
                    return new Bitmap(Properties.Resources.ExcelToDatabase);

                case "button_DatabaseToExcel":
                    return new Bitmap(Properties.Resources.DatabaseToExcel);
            }

            return null;
        }

        public DataTable GetFields(string controlId) {
            string sqlStatement = "";

            switch (controlId) {
                case "comboBox_HostName":
                    sqlStatement = @"SELECT DISTINCT ""host_name""
                                    FROM ""edb_saved_excel""
                                    ORDER BY ""host_name"";";
                    break;
                case "comboBox_BookName":
                    sqlStatement = @"SELECT DISTINCT ""book_name""
                                    FROM ""edb_saved_excel""
                                    WHERE ""host_name"" = '{0}'
                                    ORDER BY ""book_name"";".FormatWithArgs(ribbonHostName);
                    break;
                case "comboBox_SheetName":
                    sqlStatement = @"SELECT DISTINCT ""sheet_name""
                                    FROM ""edb_saved_excel""
                                    WHERE ""host_name"" = '{0}'
                                        AND ""book_name"" = '{1}'
                                    ORDER BY ""sheet_name"";".FormatWithArgs(ribbonHostName, ribbonBookName);
                    break;
                case "comboBox_ModifiedDatetime":
                    sqlStatement = @"SELECT DISTINCT CONVERT(varchar(19), ""modified_datetime"", 120)
                                    FROM ""edb_saved_excel""
                                    WHERE ""host_name"" = '{0}'
                                        AND ""book_name"" = '{1}'
                                        AND ""sheet_name"" = '{2}'
                                    ORDER BY CONVERT(varchar(19), ""modified_datetime"", 120);".FormatWithArgs(ribbonHostName, ribbonBookName, ribbonSheetName);
                    break;
            }


            DataTable result = Database.GetDataTableFromSQL(sqlStatement);

            return result;
        }
    }
}