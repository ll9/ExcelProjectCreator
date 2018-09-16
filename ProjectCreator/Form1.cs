using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectCreator
{
    public partial class Form1 : Form
    {
        public string ProjectPath { get; set; }

        public Form1()
        {
            InitializeComponent();

            InitializeProjectPath();
        }

        private void InitializeProjectPath()
        {
            if (String.IsNullOrEmpty(Properties.Settings.Default["DefaultProjectPath"].ToString()))
            {
                SelectProjectFolder();
            }
            else
            {
                ProjectPath = Properties.Settings.Default["DefaultProjectPath"].ToString();
                Debug.WriteLine(Properties.Settings.Default["DefaultProjectPath"].ToString());
            }
            System.IO.Directory.CreateDirectory($"{ProjectPath}\\Db");
        }

        private void SelectProjectFolder()
        {
            var dialog = new FolderBrowserDialog();
            dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var result = dialog.ShowDialog();

            Debug.WriteLine(result);
            Debug.WriteLine(dialog.SelectedPath);

            if (result == DialogResult.OK)
            {
                Properties.Settings.Default["DefaultProjectPath"] = dialog.SelectedPath;
                Properties.Settings.Default.Save();
                ProjectPath = dialog.SelectedPath;
            }

        }

        private void ReadExcelToDb()
        {

        }

        private void ImportExcel()
        {
            var dialog = new OpenFileDialog();
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Filter = "Excel (*.xlsx) | *.xlsx";
            var result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                Debug.WriteLine(dialog.FileName);
                System.IO.File.Copy(dialog.FileName, $"{ProjectPath}\\Db\\db.xlsx");
            }
        }

        private void openProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectProjectFolder();
        }

        private void importExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportExcel();
        }
    }
}
