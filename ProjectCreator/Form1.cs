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
        public Form1()
        {
            InitializeComponent();

            if (String.IsNullOrEmpty(Properties.Settings.Default["DefaultProjectPath"].ToString()))
            {
                SelectProjectFolder();
            }
            else
            {
                Debug.WriteLine(Properties.Settings.Default["DefaultProjectPath"].ToString());
            }
        }

        private static void SelectProjectFolder()
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
            }

        }

        private void openProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectProjectFolder();
        }
    }
}
