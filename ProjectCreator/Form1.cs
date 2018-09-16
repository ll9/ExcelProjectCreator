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
        public string ProjectPath { get; set; }

        public Form1()
        {
            InitializeComponent();

            InitializeProjectPath();
            if (File.Exists($"{ProjectPath}\\Db\\db.xlsx"))
            {
                ReadExcelToDb();
            }
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
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (XLWorkbook workBook = new XLWorkbook($"{ProjectPath}\\Db\\db.xlsx"))
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

        private void ImportExcel()
        {
            var dialog = new OpenFileDialog();
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Filter = "Excel (*.xlsx) | *.xlsx";
            var result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                Debug.WriteLine(dialog.FileName);
                System.IO.File.Copy(dialog.FileName, $"{ProjectPath}\\Db\\db.xlsx", true);
            }
            ReadExcelToDb();
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
