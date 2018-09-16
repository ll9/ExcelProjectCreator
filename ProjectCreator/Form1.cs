﻿using ClosedXML.Excel;
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

        private void ReadExcelToDb()
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (XLWorkbook workBook = new XLWorkbook(DbPath))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }
                excelGrid.DataSource = dt;
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
                ReadExcelToDb();
            }
        }

        private void ImportExcel()
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
                ReadExcelToDb();
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
