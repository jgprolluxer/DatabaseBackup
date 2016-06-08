using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                var path = $@"{args[0]} {args[1]} -p {args[2]} > {args[3]}.sql";
                var p = new Process { StartInfo = { FileName = path } };
                p.Start();
                eventLog.WriteEntry("Database backup - success.", EventLogEntryType.SuccessAudit);
            }
            catch (Exception e)
            {
                eventLog.WriteEntry($"Database backup - error - {e.Message}.", EventLogEntryType.Error);
            }
        }
    }
}
