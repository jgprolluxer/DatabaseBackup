﻿using System;
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

        public static void Database(string[] args)
        {
            if (!EventLog.SourceExists("DatabaseBackupLogSource"))
                EventLog.CreateEventSource("DatabaseBackupLogSource", "DatabaseBackupLog");

            var eventLog = new EventLog("DatabaseBackupLogSource")
            {
                Source = "DatabaseBackupLogSource",
                Log = "DatabaseBackupLog"
            };

            try
            {
                var alias = args[0];
                var exportType = (ExportType)Convert.ToInt32(args[1]);
                var cs = ConfigurationManager.ConnectionStrings[alias].ToString();
                var targetFolder = ConfigurationManager.AppSettings["TargetFolder"];
                var ds = GetDataSet(cs);

                switch (exportType)
                {
                    case ExportType.Excel:
                        ToExcel(ds, targetFolder, alias);
                        break;
                    case ExportType.Csv:
                        ToCsv(ds, targetFolder, alias);
                        break;
                    case ExportType.Sql:
                        ToSql(targetFolder, alias);
                        break;
                    case ExportType.Access:
                        ToAccess(ds, targetFolder, alias);
                        break;
                    case ExportType.ExcelCsv:
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        break;
                    case ExportType.ExcelSql:
                        ToExcel(ds, targetFolder, alias);
                        ToSql(targetFolder, alias);
                        break;
                    case ExportType.ExcelAccess:
                        ToExcel(ds, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias);
                        break;
                    case ExportType.CsvSql:
                        ToCsv(ds, targetFolder, alias);
                        ToSql(targetFolder, alias);
                        break;
                    case ExportType.CsvAccess:
                        ToCsv(ds, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias);
                        break;
                    case ExportType.SqlAccess:
                        ToSql(targetFolder, alias);
                        ToAccess(ds, targetFolder, alias);
                        break;
                    case ExportType.ExcelCsvSql:
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        ToSql(targetFolder, alias);
                        break;
                    case ExportType.ExcelCsvAccess:
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        ToAccess(ds, targetFolder, alias);
                        break;
                    case ExportType.ExcelSqlAccess:
                        ToExcel(ds, targetFolder, alias);
                        ToSql(targetFolder, alias);
                        ToAccess(ds, targetFolder, alias);
                        break;
                    case ExportType.CsvSqlAccess:
                        ToCsv(ds, targetFolder, alias);
                        ToSql(targetFolder, alias);
                        ToAccess(ds, targetFolder, alias);
                        break;
                    case ExportType.ExcelCsvSqlAccess:
                        ToExcel(ds, targetFolder, alias);
                        ToCsv(ds, targetFolder, alias);
                        ToSql(targetFolder, alias);
                        ToAccess(ds, targetFolder, alias);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                eventLog.WriteEntry("Database backup - success.", EventLogEntryType.SuccessAudit);
            }
            catch (Exception e)
            {
                eventLog.WriteEntry(string.Format("Database backup - error - {0}.", e.Message), EventLogEntryType.Error);
            }
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

        private static void ToAccess(DataSet ds, string targetFolder, string alias)
        {
            var filePath = GetFilePath(targetFolder, alias, ExportType.Access);
            var cs = ConfigurationManager.ConnectionStrings["AccessFile"].ToString().Replace("|FilePath|", filePath);
            DatasetToJet.CopyDatasetSchemaToJetDb(cs, ds, filePath);
        }

        private static void ToSql(string targetFolder, string alias)
        {
            var filePath = GetFilePath(targetFolder, alias, ExportType.Sql);

            //GET INFO FOR MYSQLDUMP
            var host = ConfigurationManager.AppSettings["server"];
            var port = ConfigurationManager.AppSettings["port"];
            var databases = ConfigurationManager.AppSettings["databases"];
            var user = ConfigurationManager.AppSettings["user"];
            var password = ConfigurationManager.AppSettings["password"];

            var sw = new StreamWriter(filePath, true);

            var process = new ProcessStartInfo();
            var command = string.Format(@"-e -P{0} -h{1} {2} -u{3} -p{4}", port, host, databases, user, password);
            process.FileName = "C:/Program Files/MySQL/MySQL Server 5.7/bin/mysqldump.exe";
            process.RedirectStandardInput = false;
            process.RedirectStandardOutput = true;
            process.Arguments = command;
            process.UseShellExecute = false;
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
