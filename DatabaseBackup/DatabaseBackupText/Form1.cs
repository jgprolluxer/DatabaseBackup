using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using DatabaseBackup;
using System.Configuration;
using System.Linq;

namespace DatabaseBackupText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var alias = ((DataRowView)comboBox1.SelectedItem).Row[0].ToString();
            var exportType = ((DataRowView)comboBox2.SelectedItem).Row[0].ToString();

            var args = new[]
            {
                alias,
                exportType
            };

            Backup.Database(args);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Description");

            foreach (var alias in ConfigurationManager.ConnectionStrings.Cast<ConnectionStringSettings>().Where(alias => alias.Name != "LocalSqlServer" && alias.Name != "AccessFile"))
            {
                var r = dt.NewRow();
                r["Id"] = alias.Name;
                r["Description"] = alias.Name;
                dt.Rows.Add(r);
            }

            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "Id";
            comboBox1.DisplayMember = "Description";
            var list = Enum.GetValues(typeof(Backup.ExportType)).Cast<Backup.ExportType>();
            //var list = new List<string>
            //{
            //    "Excel",
            //    "Csv",
            //    "Sql",
            //    "Access",
            //    "ExcelCsv",
            //    "ExcelSql",
            //    "ExcelAccess",
            //    "CsvSql",
            //    "CsvAccess",
            //    "SqlAccess"
            //};

            var dt1 = new DataTable();
            dt1.Columns.Add("Id");
            dt1.Columns.Add("Description");
            var i = 0;

            foreach (var item in list)
            {
                var r = dt1.NewRow();
                r["Id"] = i++;
                r["Description"] = item;
                dt1.Rows.Add(r);
            }

            comboBox2.DataSource = dt1;
            comboBox2.ValueMember = "Id";
            comboBox2.DisplayMember = "Description";
        }
    }
}
