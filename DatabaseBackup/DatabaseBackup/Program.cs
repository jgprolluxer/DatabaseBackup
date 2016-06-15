using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace DatabaseBackup
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            if (!EventLog.SourceExists("DatabaseBackupLogSource"))
                EventLog.CreateEventSource("DatabaseBackupLogSource", "DatabaseBackupLog");

            var eventLog = new EventLog("DatabaseBackupLogSource")
            {
                Source = "DatabaseBackupLogSource",
                Log = "DatabaseBackupLog"
            };

            var errList = new List<string>();

            try
            {
                var b = Backup.Database(args, ref errList);

                if (b)
                {
                    eventLog.WriteEntry("Database backup - success.", EventLogEntryType.SuccessAudit);
                }
                else
                {
                    eventLog.WriteEntry(string.Format("Database backup - error - {0}", string.Join(",", errList)), EventLogEntryType.Error);
                }
            }
            catch (Exception e)
            {
                eventLog.WriteEntry(string.Format("Database backup - error - {0} {1} {2}", e.Message, e.StackTrace, string.Join(",", errList)), EventLogEntryType.Error);
            }
        }
    }
}
