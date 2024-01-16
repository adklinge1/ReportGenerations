using System.Drawing;

namespace WindowsFormsApp1
{
    partial class ReportGeneratorForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnGenerateDocument = new System.Windows.Forms.Button();
            this.txtDirectoryPath = new System.Windows.Forms.TextBox();
            this.folderPathTxtLabel = new System.Windows.Forms.Label();
            this.templatePathTextBox = new System.Windows.Forms.TextBox();
            this.templateDocumentLabel = new System.Windows.Forms.Label();
            this.reportNameTextBox = new System.Windows.Forms.TextBox();
            this.reportNameLabel = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGenerateDocument
            // 
            this.btnGenerateDocument.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.btnGenerateDocument.Location = new System.Drawing.Point(266, 284);
            this.btnGenerateDocument.Name = "btnGenerateDocument";
            this.btnGenerateDocument.Size = new System.Drawing.Size(270, 117);
            this.btnGenerateDocument.TabIndex = 0;
            this.btnGenerateDocument.Text = "Generate Report";
            this.btnGenerateDocument.UseVisualStyleBackColor = true;
            this.btnGenerateDocument.Click += new System.EventHandler(this.btnGenerateDocument_Click);
            // 
            // txtDirectoryPath
            // 
            this.txtDirectoryPath.Location = new System.Drawing.Point(21, 48);
            this.txtDirectoryPath.Name = "txtDirectoryPath";
            this.txtDirectoryPath.Size = new System.Drawing.Size(326, 20);
            this.txtDirectoryPath.TabIndex = 1;
            this.txtDirectoryPath.Text = "C:\\";
            this.txtDirectoryPath.TextChanged += new System.EventHandler(this.txtDirectoryPath_TextChanged);
            // 
            // folderPathTxtLabel
            // 
            this.folderPathTxtLabel.AutoSize = true;
            this.folderPathTxtLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.folderPathTxtLabel.Location = new System.Drawing.Point(18, 18);
            this.folderPathTxtLabel.Name = "folderPathTxtLabel";
            this.folderPathTxtLabel.Size = new System.Drawing.Size(283, 17);
            this.folderPathTxtLabel.TabIndex = 2;
            this.folderPathTxtLabel.Text = "Insert the path to the folder of images";
            this.folderPathTxtLabel.Click += new System.EventHandler(this.label1_Click);
            // 
            // templatePathTextBox
            // 
            this.templatePathTextBox.Location = new System.Drawing.Point(21, 181);
            this.templatePathTextBox.Name = "templatePathTextBox";
            this.templatePathTextBox.Size = new System.Drawing.Size(319, 20);
            this.templatePathTextBox.TabIndex = 3;
            // 
            // templateDocumentLabel
            // 
            this.templateDocumentLabel.AutoSize = true;
            this.templateDocumentLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.templateDocumentLabel.Location = new System.Drawing.Point(18, 153);
            this.templateDocumentLabel.Name = "templateDocumentLabel";
            this.templateDocumentLabel.Size = new System.Drawing.Size(268, 17);
            this.templateDocumentLabel.TabIndex = 4;
            this.templateDocumentLabel.Text = "Word Template document (optional)";
            // 
            // reportNameTextBox
            // 
            this.reportNameTextBox.Location = new System.Drawing.Point(21, 113);
            this.reportNameTextBox.Name = "reportNameTextBox";
            this.reportNameTextBox.Size = new System.Drawing.Size(164, 20);
            this.reportNameTextBox.TabIndex = 5;
            // 
            // reportNameLabel
            // 
            this.reportNameLabel.AutoSize = true;
            this.reportNameLabel.Location = new System.Drawing.Point(18, 87);
            this.reportNameLabel.Name = "reportNameLabel";
            this.reportNameLabel.Size = new System.Drawing.Size(70, 13);
            this.reportNameLabel.TabIndex = 6;
            this.reportNameLabel.Text = "Report Name";
            this.reportNameLabel.Font = new System.Drawing.Font("Boucherie Block", 10F, System.Drawing.FontStyle.Bold);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 428);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(800, 22);
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // ReportGeneratorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.reportNameLabel);
            this.Controls.Add(this.reportNameTextBox);
            this.Controls.Add(this.templateDocumentLabel);
            this.Controls.Add(this.templatePathTextBox);
            this.Controls.Add(this.folderPathTxtLabel);
            this.Controls.Add(this.txtDirectoryPath);
            this.Controls.Add(this.btnGenerateDocument);
            this.Name = "ReportGeneratorForm";
            this.Text = "Report Generator";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerateDocument;
        private System.Windows.Forms.TextBox txtDirectoryPath ;
        private System.Windows.Forms.Label folderPathTxtLabel;
        private System.Windows.Forms.TextBox templatePathTextBox;
        private System.Windows.Forms.Label templateDocumentLabel;

        private System.Windows.Forms.TextBox reportNameTextBox;
        private System.Windows.Forms.Label reportNameLabel;
        
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
    }
}

