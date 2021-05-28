﻿namespace BulkPDF
{
    partial class FieldOptionForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FieldOptionForm));
            this.cbDataSourceColumns = new System.Windows.Forms.ComboBox();
            this.bSet = new System.Windows.Forms.Button();
            this.cbUseValueFromDataSource = new System.Windows.Forms.CheckBox();
            this.cbReadOnly = new System.Windows.Forms.CheckBox();
            this.cbRow = new System.Windows.Forms.NumericUpDown();
            this.labelRow = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.cbRow)).BeginInit();
            this.SuspendLayout();
            // 
            // cbDataSourceColumns
            // 
            this.cbDataSourceColumns.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbDataSourceColumns.FormattingEnabled = true;
            resources.ApplyResources(this.cbDataSourceColumns, "cbDataSourceColumns");
            this.cbDataSourceColumns.Name = "cbDataSourceColumns";
            this.cbDataSourceColumns.SelectedIndexChanged += new System.EventHandler(this.cbDataSourceColumns_SelectedIndexChanged);
            // 
            // bSet
            // 
            resources.ApplyResources(this.bSet, "bSet");
            this.bSet.Name = "bSet";
            this.bSet.UseVisualStyleBackColor = true;
            this.bSet.Click += new System.EventHandler(this.bSet_Click);
            // 
            // cbUseValueFromDataSource
            // 
            resources.ApplyResources(this.cbUseValueFromDataSource, "cbUseValueFromDataSource");
            this.cbUseValueFromDataSource.Name = "cbUseValueFromDataSource";
            this.cbUseValueFromDataSource.UseVisualStyleBackColor = true;
            this.cbUseValueFromDataSource.CheckedChanged += new System.EventHandler(this.cbUseValueFromDataSource_CheckedChanged);
            // 
            // cbReadOnly
            // 
            resources.ApplyResources(this.cbReadOnly, "cbReadOnly");
            this.cbReadOnly.Name = "cbReadOnly";
            this.cbReadOnly.UseVisualStyleBackColor = true;
            this.cbReadOnly.CheckedChanged += new System.EventHandler(this.cbReadOnly_CheckedChanged);
            // 
            // cbRow
            // 
            resources.ApplyResources(this.cbRow, "cbRow");
            this.cbRow.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.cbRow.Name = "cbRow";
            this.cbRow.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.cbRow.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // labelRow
            // 
            resources.ApplyResources(this.labelRow, "labelRow");
            this.labelRow.Name = "labelRow";
            // 
            // FieldOptionForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelRow);
            this.Controls.Add(this.cbRow);
            this.Controls.Add(this.cbReadOnly);
            this.Controls.Add(this.cbUseValueFromDataSource);
            this.Controls.Add(this.bSet);
            this.Controls.Add(this.cbDataSourceColumns);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "FieldOptionForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FieldOptionForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.cbRow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbDataSourceColumns;
        private System.Windows.Forms.Button bSet;
        private System.Windows.Forms.CheckBox cbUseValueFromDataSource;
        private System.Windows.Forms.CheckBox cbReadOnly;
        private System.Windows.Forms.NumericUpDown cbRow;
        private System.Windows.Forms.Label labelRow;
    }
}