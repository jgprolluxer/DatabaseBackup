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

namespace DatabaseBackup
{
    public class Backup
    {
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
                var ds = GetDataSet(args[0]);
                ExportDataSet(ds, args[1]);
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

        private static void ExportDataSet(DataSet ds, string destination)
        {
            var folder = Path.GetDirectoryName(destination);
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);

            using (var workbook = SpreadsheetDocument.Create(destination, SpreadsheetDocumentType.Workbook))
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
    }
}
