﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectCreator
{
    public partial class Form1 : Form
    {
        public readonly string DB_FOLDER = "Db";
        public readonly string DB_NAME = "db.xlsx";

        public string ProjectPath => Properties.Settings.Default["DefaultProjectPath"].ToString();
        public bool ProjectPathExists => !string.IsNullOrEmpty(ProjectPath);
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
                InitDb();
            }

        }

        /// <summary>
        /// Loads Db to Datagrid If it exists and
        /// Enables Excel Menu Item
        /// </summary>
        private void InitDb()
        {
            importExcelToolStripMenuItem.Enabled = true;
            Directory.CreateDirectory(DbFolderPath);

            if (DbExists)
            {
                excelGrid.DataSource = ExcelHandler.GetDataTable(DbPath);
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
                Debug.WriteLine(dialog.FileName);
                File.Copy(dialog.FileName, DbPath, true);
                excelGrid.DataSource = ExcelHandler.GetDataTable(DbPath);
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
    }
}
