using System;
using System.Drawing;
using System.Windows.Forms;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportGeneratorForm));
            this.btnGenerateDocument = new System.Windows.Forms.Button();
            this.txtDirectoryPath = new System.Windows.Forms.TextBox();
            this.statusTxtLabel = new System.Windows.Forms.Label();
            this.templatePathTextBox = new System.Windows.Forms.TextBox();
            this.reportNameTextBox = new System.Windows.Forms.TextBox();
            this.reportNameLabel = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.excelFilePahTextBox = new System.Windows.Forms.TextBox();
            this.browseImgFolderButton = new System.Windows.Forms.Button();
            this.browseExcelFileButton = new System.Windows.Forms.Button();
            this.browseTemplateButton = new System.Windows.Forms.Button();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGenerateDocument
            // 
            this.btnGenerateDocument.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.btnGenerateDocument.Location = new System.Drawing.Point(448, 665);
            this.btnGenerateDocument.Margin = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.btnGenerateDocument.Name = "btnGenerateDocument";
            this.btnGenerateDocument.Size = new System.Drawing.Size(390, 131);
            this.btnGenerateDocument.TabIndex = 0;
            this.btnGenerateDocument.Text = "Generate Report";
            this.btnGenerateDocument.UseVisualStyleBackColor = true;
            this.btnGenerateDocument.Click += new System.EventHandler(this.btnGenerateDocument_Click);
            // 
            // txtDirectoryPath
            // 
            this.txtDirectoryPath.Location = new System.Drawing.Point(448, 160);
            this.txtDirectoryPath.Margin = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.txtDirectoryPath.Multiline = true;
            this.txtDirectoryPath.Name = "txtDirectoryPath";
            this.txtDirectoryPath.Size = new System.Drawing.Size(863, 92);
            this.txtDirectoryPath.TabIndex = 1;
            // 
            // statusTxtLabel
            // 
            this.statusTxtLabel.AutoSize = true;
            this.statusTxtLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.statusTxtLabel.Location = new System.Drawing.Point(1412, 940);
            this.statusTxtLabel.Margin = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.statusTxtLabel.Name = "statusTxtLabel";
            this.statusTxtLabel.Size = new System.Drawing.Size(0, 39);
            this.statusTxtLabel.TabIndex = 2;
            // 
            // templatePathTextBox
            // 
            this.templatePathTextBox.Location = new System.Drawing.Point(448, 491);
            this.templatePathTextBox.Margin = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.templatePathTextBox.Multiline = true;
            this.templatePathTextBox.Name = "templatePathTextBox";
            this.templatePathTextBox.Size = new System.Drawing.Size(863, 92);
            this.templatePathTextBox.TabIndex = 3;
            // 
            // reportNameTextBox
            // 
            this.reportNameTextBox.Location = new System.Drawing.Point(448, 38);
            this.reportNameTextBox.Margin = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.reportNameTextBox.Multiline = true;
            this.reportNameTextBox.Name = "reportNameTextBox";
            this.reportNameTextBox.Size = new System.Drawing.Size(431, 85);
            this.reportNameTextBox.TabIndex = 5;
            this.reportNameTextBox.TextChanged += new System.EventHandler(this.reportNameTextBox_TextChanged);
            // 
            // reportNameLabel
            // 
            this.reportNameLabel.AutoSize = true;
            this.reportNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.reportNameLabel.Location = new System.Drawing.Point(89, 54);
            this.reportNameLabel.Margin = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.reportNameLabel.Name = "reportNameLabel";
            this.reportNameLabel.Size = new System.Drawing.Size(230, 39);
            this.reportNameLabel.TabIndex = 6;
            this.reportNameLabel.Text = "Report Name";
            this.reportNameLabel.Click += new System.EventHandler(this.reportNameLabel_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(40, 40);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 1028);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(3, 0, 37, 0);
            this.statusStrip1.Size = new System.Drawing.Size(2526, 30);
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            this.statusStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.statusStrip1_ItemClicked);
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(267, 14);
            // 
            // excelFilePahTextBox
            // 
            this.excelFilePahTextBox.Location = new System.Drawing.Point(448, 325);
            this.excelFilePahTextBox.Margin = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.excelFilePahTextBox.Multiline = true;
            this.excelFilePahTextBox.Name = "excelFilePahTextBox";
            this.excelFilePahTextBox.Size = new System.Drawing.Size(863, 92);
            this.excelFilePahTextBox.TabIndex = 9;
            // 
            // browseImgFolderButton
            // 
            this.browseImgFolderButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.browseImgFolderButton.Location = new System.Drawing.Point(51, 160);
            this.browseImgFolderButton.Name = "browseImgFolderButton";
            this.browseImgFolderButton.Size = new System.Drawing.Size(337, 92);
            this.browseImgFolderButton.TabIndex = 10;
            this.browseImgFolderButton.Text = "Select Images Folder";
            this.browseImgFolderButton.UseVisualStyleBackColor = true;
            this.browseImgFolderButton.Click += new System.EventHandler(this.browseImgFolder_Click);
            // 
            // browseExcelFileButton
            // 
            this.browseExcelFileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.browseExcelFileButton.Location = new System.Drawing.Point(60, 325);
            this.browseExcelFileButton.Name = "browseExcelFileButton";
            this.browseExcelFileButton.Size = new System.Drawing.Size(337, 92);
            this.browseExcelFileButton.TabIndex = 11;
            this.browseExcelFileButton.Text = "Select Excel File";
            this.browseExcelFileButton.UseVisualStyleBackColor = true;
            this.browseExcelFileButton.Click += new System.EventHandler(this.selectExcelFile_Click);
            // 
            // browseTemplateButton
            // 
            this.browseTemplateButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.browseTemplateButton.Location = new System.Drawing.Point(60, 491);
            this.browseTemplateButton.Name = "browseTemplateButton";
            this.browseTemplateButton.Size = new System.Drawing.Size(337, 92);
            this.browseTemplateButton.TabIndex = 12;
            this.browseTemplateButton.Text = "Select Template File (optional)";
            this.browseTemplateButton.UseVisualStyleBackColor = true;
            this.browseTemplateButton.Click += new System.EventHandler(this.templateBrowser_Click);
            // 
            // ReportGeneratorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(16F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(2526, 1058);
            this.Controls.Add(this.browseTemplateButton);
            this.Controls.Add(this.browseExcelFileButton);
            this.Controls.Add(this.browseImgFolderButton);
            this.Controls.Add(this.excelFilePahTextBox);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.reportNameLabel);
            this.Controls.Add(this.reportNameTextBox);
            this.Controls.Add(this.templatePathTextBox);
            this.Controls.Add(this.statusTxtLabel);
            this.Controls.Add(this.txtDirectoryPath);
            this.Controls.Add(this.btnGenerateDocument);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.Name = "ReportGeneratorForm";
            this.Text = "Report Generator";
            this.Load += new System.EventHandler(this.ReportGeneratorForm_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerateDocument;
        private System.Windows.Forms.TextBox txtDirectoryPath ;
        private System.Windows.Forms.Label statusTxtLabel;
        private System.Windows.Forms.TextBox templatePathTextBox;

        private System.Windows.Forms.TextBox reportNameTextBox;
        private System.Windows.Forms.Label reportNameLabel;
        
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.TextBox excelFilePahTextBox;
        private System.Windows.Forms.Button browseImgFolderButton;
        private System.Windows.Forms.Button browseExcelFileButton;
        private Button browseTemplateButton;
    }
}

