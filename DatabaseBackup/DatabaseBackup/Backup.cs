using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;
using System.IO;
using System.Configuration;
using System.Text;
using ADOXJetXML;

namespace DatabaseBackup
{
    public static class Backup
    {
        public enum ExportType
        {
            Excel = 0,
            Csv,
            Sql,
            Access,
            ExcelCsv,
            ExcelSql,
            ExcelAccess,
            CsvSql,
            CsvAccess,
            SqlAccess,
            ExcelCsvSql,
            ExcelCsvAccess,
            ExcelSqlAccess,
            CsvSqlAccess,
            ExcelCsvSqlAccess
        }

        public static bool Database(string[] args, ref List<string> errList)
        {
            var err = "";

            if (args.Length == 2)
            {
                var alias = args[0];
                var exportType = (ExportType) Convert.ToInt32(args[1]);
                var cs = ConfigurationManager.ConnectionStrings[alias].ConnectionString;
                var targetFolder = ConfigurationManager.AppSettings["TargetFolder"];
                DataSet ds;

                switch (exportType)
                {
                    case ExportType.Excel:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        break;
                    case ExportType.Csv:
                        ds = GetDataSet(cs);
                        ToCsv(ds, targetFolder, alias);
                        break;
                    case ExportType.Sql:
                        ToSql(cs, targetFolder, alias);
                        break;
                    case ExportType.Access:
                        ds = GetDataSet(cs);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    case ExportType.ExcelCsv:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        break;
                    case ExportType.ExcelSql:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        ToSql(cs, targetFolder, alias);
                        break;
                    case ExportType.ExcelAccess:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    case ExportType.CsvSql:
                        ds = GetDataSet(cs);
                        ToCsv(ds, targetFolder, alias);
                        ToSql(cs, targetFolder, alias);
                        break;
                    case ExportType.CsvAccess:
                        ds = GetDataSet(cs);
                        ToCsv(ds, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    case ExportType.SqlAccess:
                        ds = GetDataSet(cs);
                        ToSql(cs, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    case ExportType.ExcelCsvSql:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        ToSql(cs, targetFolder, alias);
                        break;
                    case ExportType.ExcelCsvAccess:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    case ExportType.ExcelSqlAccess:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        ToSql(cs, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    case ExportType.CsvSqlAccess:
                        ds = GetDataSet(cs);
                        ToCsv(ds, targetFolder, alias);
                        ToSql(cs, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    case ExportType.ExcelCsvSqlAccess:
                        ds = GetDataSet(cs);
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        ToSql(cs, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias, ref err);
                        break;
                    default:
                        errList.Add(string.Format("Wrong arguments: {0} {1}", args[0], args[1]));
                        break;
                }
            }
            else
            {
                errList.Add("Wrong argument length: " + args.Length);
            }

            if(err != "")
                errList.Add(err);

            return errList.Count == 0;
        }

        private static DataSet GetDataSet(string connString)
        {
            var ds = new DataSet();

            using (var conn = new MySqlConnection(connString))
            {
                conn.Open();
                var tables = conn.GetSchema("Tables", new[] { null, null, null, "BASE TABLE" });

                foreach (DataRow table in tables.Rows)
                {
                    var tableName = table["TABLE_NAME"].ToString();

                    var daAuthors = new MySqlDataAdapter(string.Format("Select * From {0}", tableName), conn)
                    {
                        MissingSchemaAction = MissingSchemaAction.AddWithKey
                    };

                    daAuthors.FillSchema(ds, SchemaType.Source, tableName);
                    daAuthors.Fill(ds, tableName);
                }
            }

            return ds;
        }

        private static void ToExcel(DataSet ds, string targetFolder, string alias)
        {
            var filePath = GetFilePath(targetFolder, alias, ExportType.Excel);

            using (var workbook = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook { Sheets = new Sheets() };

                foreach (DataTable table in ds.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    uint sheetId = 1;
                    if (sheets.Elements<Sheet>().Any())
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    var sheetList = new List<Sheet>
                    {
                        new Sheet {Id = relationshipId, SheetId = sheetId, Name = table.TableName}
                    };

                    sheets.Append(sheetList);

                    var headerRow = new Row();

                    var columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.ColumnName)
                        };
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        var newRow = new Row();
                        var r = dsrow;
                        foreach (var cell in columns.Select(col => new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(r[col].ToString())
                        }))
                        {
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
            }
        }

        private static string GetFilePath(string targetFolder, string alias, ExportType exportType)
        {
            var folder = Path.GetDirectoryName(targetFolder);
            folder = string.Format(@"{0}\{1}", folder, DateTime.Now.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            var extension = "";

            switch (exportType)
            {
                case ExportType.Excel:
                    extension = "xlsx";
                    break;
                case ExportType.Csv:
                    extension = "csv";
                    break;
                case ExportType.Access:
                    extension = "accdb";
                    break;
                case ExportType.Sql:
                    extension = "sql";
                    break;
            }

            return string.Format(@"{0}\{1}.{2}", folder, alias, extension);
        }

        private static void ToCsv(DataSet ds, string targetFolder, string alias)
        {
            foreach (DataTable dt in ds.Tables)
            {
                var result = new StringBuilder();

                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    result.Append(dt.Columns[i].ColumnName);
                    result.Append(i == dt.Columns.Count - 1 ? "\n" : ",");
                }

                foreach (DataRow row in dt.Rows)
                {
                    for (var i = 0; i < dt.Columns.Count; i++)
                    {
                        result.Append(row[i]);
                        result.Append(i == dt.Columns.Count - 1 ? "\n" : ",");
                    }
                }

                var fileName = string.Format("{0}-{1}", alias, dt.TableName);
                var filePath = GetFilePath(targetFolder, fileName, ExportType.Csv);
                File.WriteAllText(filePath, result.ToString(), Encoding.Default);
            }
        }

        private static void ToAccess(DataSet ds, string targetFolder, string alias, ref string err)
        {
            var filePath = GetFilePath(targetFolder, alias, ExportType.Access);
            var cs = ConfigurationManager.ConnectionStrings["AccessFile"].ToString().Replace("|FilePath|", filePath);
            DatasetToJet.CopyDatasetSchemaToJetDb(cs, ds, filePath, ref err);
        }

        private static void ToSql(string connectionString, string targetFolder, string alias)
        {
            var filePath = GetFilePath(targetFolder, alias, ExportType.Sql);
            //GET INFO FOR MYSQLDUMP
            var b = new MySqlConnectionStringBuilder(connectionString);
            var sw = new StreamWriter(filePath, true);

            var process = new ProcessStartInfo();
            var command = string.Format(@"-e -P{0} -h{1} {2} -u{3} -p{4}", b.Port, b.Server, b.Database, b.UserID, b.Password);
            process.FileName = ConfigurationManager.AppSettings["MySqlDump"];
            process.RedirectStandardInput = false;
            process.RedirectStandardOutput = true;
            process.Arguments = command;
            process.UseShellExecute = false;
            process.WindowStyle = ProcessWindowStyle.Hidden;
            process.CreateNoWindow = true;
            var proc = Process.Start(process);

            if (proc != null)
            {
                var response = proc.StandardOutput.ReadToEnd();

                //SAVE RESPONSE IN SQL FILE
                sw.WriteLine(response);
                proc.WaitForExit();
            }

            sw.Close();
        }
    }
}
