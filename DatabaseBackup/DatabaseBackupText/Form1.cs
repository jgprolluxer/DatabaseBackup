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
            var args = new[]
            {
                textBox1.Text,
                textBox2.Text,
                textBox3.Text,
                textBox4.Text
            };

            Backup.Database(args);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
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
