using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;

namespace ProjectCreator
{
    public partial class Form1 : Form
    {
        public readonly string DB_FOLDER = "Db";
        public readonly string DB_NAME = "db.xlsx";

        public string ProjectPath => Properties.Settings.Default["DefaultProjectPath"].ToString();
        public bool ProjectPathExists => !string.IsNullOrEmpty(ProjectPath);
        public string BackupsFolderPath => $"{ProjectPath}\\Backups";
        public string DbFolderPath
        {
            get
            {
                if (!ProjectPathExists)
                {
                    throw new InvalidOperationException("Cannot Get Db Path when Project Path does not exist");
                }
                return $"{ProjectPath}\\{DB_FOLDER}";
            }
        }
        private string DbPath => $"{DbFolderPath}\\{DB_NAME}";
        public bool DbExists => File.Exists(DbPath);


        public Form1()
        {
            InitializeComponent();

            InitializeProject();
        }

        private void InitializeProject()
        {
            if (!ProjectPathExists)
            {
                importExcelToolStripMenuItem.Enabled = false;
                restoreBackupToolStripMenuItem.Enabled = false;
                createBackupToolStripMenuItem.Enabled = false;
            }
            else
            {
                InitDb();
            }
        }

        private void SelectProjectFolder()
        {
            var dialog = new FolderBrowserDialog
            {
                SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            var result = dialog.ShowDialog();


            if (result == DialogResult.OK)
            {
                Properties.Settings.Default["DefaultProjectPath"] = dialog.SelectedPath;
                Properties.Settings.Default.Save();
                CreateProjectFolders();
                InitDb();
            }

        }

        private void CreateProjectFolders()
        {
            Directory.CreateDirectory(DbFolderPath);
            Directory.CreateDirectory(BackupsFolderPath);

        }

        /// <summary>
        /// Loads Db to Datagrid If it exists and
        /// Enables Excel Menu Item
        /// </summary>
        private void InitDb()
        {
            importExcelToolStripMenuItem.Enabled = true;
            restoreBackupToolStripMenuItem.Enabled = true;

            if (DbExists)
            {
                excelGrid.DataSource = ExcelHandler.GetDataTable(DbPath);
                createBackupToolStripMenuItem.Enabled = true;
            }
            else
            {
                excelGrid.DataSource = new DataTable();
                createBackupToolStripMenuItem.Enabled = false;
            }
        }

        private void OpenImportExcelDialog()
        {
            var dialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Excel (*.xlsx) | *.xlsx"
            };
            var result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                File.Copy(dialog.FileName, DbPath, true);
                excelGrid.DataSource = ExcelHandler.GetDataTable(DbPath);
                createBackupToolStripMenuItem.Enabled = true;
            }
        }

        private void openProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectProjectFolder();
        }

        private void importExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenImportExcelDialog();
        }

        private void createBackupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var zipPath = $"{BackupsFolderPath}\\backup-{DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss")}.zip";

            ZipFile.CreateFromDirectory(DbFolderPath, zipPath);

            MessageBox.Show($"Backup has been cretaed in \n{zipPath}");
        }

        private void restoreBackupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                InitialDirectory = BackupsFolderPath,
                Filter = "Zip File (*.zip) | *.zip"
            };

            var result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                var tempFolder = $"{Path.GetTempPath()}\\excelRestoreFolder";
                string tempFile = $"{tempFolder}\\{DB_NAME}";

                Directory.CreateDirectory(tempFolder);
                ZipFile.ExtractToDirectory(dialog.FileName, tempFolder);

                File.Copy(tempFile, DbPath, true);
                File.Delete(tempFile);
                Directory.Delete(tempFolder);

                excelGrid.DataSource = ExcelHandler.GetDataTable(DbPath);
                createBackupToolStripMenuItem.Enabled = true;
                MessageBox.Show("Restore Successful");
            }
        }
    }
}
