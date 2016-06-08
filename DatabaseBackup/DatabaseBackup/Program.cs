using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseBackup
{
    public class Program
    {
        static void Main(string[] args)
        {
            Backup.Database(args);
        }
    }
}
