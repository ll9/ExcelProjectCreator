namespace ProjectCreator
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openProjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelGrid = new System.Windows.Forms.DataGridView();
            this.createBackupToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.excelGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuToolStripMenuItem
            // 
            this.menuToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openProjectToolStripMenuItem,
            this.importExcelToolStripMenuItem,
            this.createBackupToolStripMenuItem});
            this.menuToolStripMenuItem.Name = "menuToolStripMenuItem";
            this.menuToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.menuToolStripMenuItem.Text = "Menu";
            // 
            // openProjectToolStripMenuItem
            // 
            this.openProjectToolStripMenuItem.Name = "openProjectToolStripMenuItem";
            this.openProjectToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.openProjectToolStripMenuItem.Text = "Open Project";
            this.openProjectToolStripMenuItem.Click += new System.EventHandler(this.openProjectToolStripMenuItem_Click);
            // 
            // importExcelToolStripMenuItem
            // 
            this.importExcelToolStripMenuItem.Name = "importExcelToolStripMenuItem";
            this.importExcelToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.importExcelToolStripMenuItem.Text = "Import Excel";
            this.importExcelToolStripMenuItem.Click += new System.EventHandler(this.importExcelToolStripMenuItem_Click);
            // 
            // excelGrid
            // 
            this.excelGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.excelGrid.Location = new System.Drawing.Point(27, 62);
            this.excelGrid.Name = "excelGrid";
            this.excelGrid.Size = new System.Drawing.Size(506, 294);
            this.excelGrid.TabIndex = 1;
            // 
            // createBackupToolStripMenuItem
            // 
            this.createBackupToolStripMenuItem.Name = "createBackupToolStripMenuItem";
            this.createBackupToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.createBackupToolStripMenuItem.Text = "Create Backup";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.excelGrid);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.excelGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openProjectToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importExcelToolStripMenuItem;
        private System.Windows.Forms.DataGridView excelGrid;
        private System.Windows.Forms.ToolStripMenuItem createBackupToolStripMenuItem;
    }
}

