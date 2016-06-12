using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;
using System.IO;
using System.Configuration;
using System.Text;
using System.Data.OleDb;

namespace DatabaseBackup
{
    public class Backup
    {
        enum ExportType
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
            SqlAccess
        }

        public static void Database(string[] args)
        {
            //string path = @"C:\Program Files\MySQL\MySQL Server 5.5\bin\mysqldump.exe -u " + txtBoxDBUsername.Text + @" -p " + txtBoxDBName.Text + @" > " + txtBoxDBName.Text + @".sql";
            if (!EventLog.SourceExists("DatabaseBackupLogSource"))
                EventLog.CreateEventSource("DatabaseBackupLogSource", "DatabaseBackupLog");

            var eventLog = new EventLog("DatabaseBackupLogSource")
            {
                Source = "DatabaseBackupLogSource",
                Log = "DatabaseBackupLog"
            };

            try
            {
                //var s = "Data Source='sql3.freesqldatabase.com';Port=3306;Database='sql386557';UID='sql386557'PWD='iK2!kM1*';";
                //excel, csv, sql, access
                var alias = args[0];
                var exportType = (ExportType)Convert.ToInt32(args[1]);
                var cs = ConfigurationManager.ConnectionStrings[alias].ToString();
                var targetFolder = ConfigurationManager.AppSettings["TargetFolder"];
                var ds = GetDataSet(cs);
                var filePath = GetFilePath(targetFolder, alias, exportType);

                switch (exportType)
                {
                    case ExportType.Excel:
                        ToExcel(ds, filePath);
                        break;
                    case ExportType.Csv:
                        ToCsv(ds, targetFolder, alias);
                        break;
                }

                //var path = $@"{args[0]} {args[1]} -p {args[2]} > {args[3]}.sql";
                //var p = new Process { StartInfo = { FileName = path } };
                //p.Start();
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

        private static void ToExcel(DataSet ds, string filePath)
        {
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
                        sheetId =
                            sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
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
                        foreach (var cell in columns.Select(col => new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(dsrow[col].ToString())
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
            }

            return string.Format(@"{0}\{1}.{2}", folder, alias, extension);
        }

        private static void ToCsv(DataSet ds, string targetFolder, string alias)
        {
            foreach (DataTable dt in ds.Tables)
            {
                var result = new StringBuilder();

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    result.Append(dt.Columns[i].ColumnName);
                    result.Append(i == dt.Columns.Count - 1 ? "\n" : ",");
                }

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        result.Append(row[i].ToString());
                        result.Append(i == dt.Columns.Count - 1 ? "\n" : ",");
                    }
                }

                //var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(result.ToString());
                //MemoryStream stream = new MemoryStream(bytes);
                //StreamReader reader = new StreamReader(stream);
                var fileName = string.Format("{0}-{1}", alias, dt.TableName);
                var filePath = GetFilePath(targetFolder, fileName, ExportType.Csv);
                File.WriteAllText(filePath, result.ToString(), Encoding.Default);
            }
        }

        private static void ToAccess(DataSet ds) {
            OleDbConnection myConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"Database.accdb\";Persist Security Info=False;");
            OleDbCommand cmd = new OleDbCommand();
            OleDbCommand cmd1 = new OleDbCommand();
            DataTable dtCSV = new DataTable();
            dtCSV = ds.Tables[0];
            cmd.Connection = myConnection;
            cmd.CommandType = CommandType.Text;
            cmd1.Connection = myConnection;
            cmd1.CommandType = CommandType.Text;

            myConnection.Open();

            foreach (DataTable dt in ds.Tables)
            {
                for (int i = 0; i <= dtCSV.Rows.Count - 1; i++)
                {
                    cmd.CommandText = "INSERT INTO " + dt.TableName + "(ID, " + dtCSV.Columns[0].ColumnName.Trim() + ") VALUES (" + (i + 1) + ",'" + dtCSV.Rows[i].ItemArray.GetValue(0) + "')";

                    cmd.ExecuteNonQuery();

                    for (int j = 1; j <= dtCSV.Columns.Count - 1; j++)
                    {
                        cmd1.CommandText = "UPDATE " + dt.TableName + " SET [" + dtCSV.Columns[j].ColumnName.Trim() + "] = '" + dtCSV.Rows[i].ItemArray.GetValue(j) + "' WHERE ID = " + (i + 1);

                        cmd1.ExecuteNonQuery();
                    }
                }
            }

            myConnection.Close();
        }
    }
}
