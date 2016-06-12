using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DatabaseBackup;
using System.Configuration;

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
            //string path = @"C:\Program Files\MySQL\MySQL Server 5.5\bin\mysqldump.exe -u " + txtBoxDBUsername.Text + @" -p " + txtBoxDBName.Text + @" > " + txtBoxDBName.Text + @".sql";
            var alias = comboBox1.SelectedItem.ToString();
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
            foreach (ConnectionStringSettings alias in ConfigurationManager.ConnectionStrings)
            {
                if (alias.Name != "LocalSqlServer")
                {
                    comboBox1.Items.Add(alias.Name);
                }
            }

            var list = new List<string>();
            list.Add("Excel");
            list.Add("Csv");
            list.Add("Sql");
            list.Add("Access");
            list.Add("ExcelCsv");
            list.Add("ExcelSql");
            list.Add("ExcelAccess");
            list.Add("CsvSql");
            list.Add("CsvAccess");
            list.Add("SqlAccess");

            var dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Description");
            var i = 0;

            foreach(var item in list) {
                var r = dt.NewRow();
                r["Id"] = i++;
                r["Description"] = item;
                dt.Rows.Add(r);
            }

            comboBox2.DataSource = dt;
            comboBox2.ValueMember = "Id";
            comboBox2.DisplayMember = "Description";
            //textBox1.Text = @"C:\Program Files\MySQL\MySQL Server 5.5\bin\mysqldump.exe -u ";
            //textBox1.Text = @"server=sql3.freesqldatabase.com;user=sql386557;database=sql386557;password=iK2!kM1*;";
            //textBox1.Text = @"Data Source='sql3.freesqldatabase.com';Port=3306;Database='sql386557';UID='sql386557'PWD='iK2!kM1*';";
            textBox1.Text = @"Data Source=sql3.freesqldatabase.com;port=3306;Initial Catalog=sql386557;User Id=sql386557;password=iK2!kM1*";
            textBox2.Text = @"c:\Development\Prollux\Test\1.xlsx";
            textBox3.Text = @"database";
            textBox4.Text = @"filename";
        }
    }
}
