namespace BulkPDF
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.bNext = new System.Windows.Forms.Button();
            this.bBack = new System.Windows.Forms.Button();
            this.bFinish = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.bDonate = new System.Windows.Forms.Button();
            this.seperatorLine1 = new BulkPDF.SeperatorLine();
            this.wizardPages = new BulkPDF.WizardPages();
            this.tpWelcome = new System.Windows.Forms.TabPage();
            this.llSupport = new System.Windows.Forms.LinkLabel();
            this.llBulkPDFde = new System.Windows.Forms.LinkLabel();
            this.llLicenses = new System.Windows.Forms.LinkLabel();
            this.rtbWelcome = new System.Windows.Forms.RichTextBox();
            this.lVersion = new System.Windows.Forms.Label();
            this.lBulkPDF = new System.Windows.Forms.Label();
            this.bLoadConfiguration = new System.Windows.Forms.Button();
            this.llDokumentation = new System.Windows.Forms.LinkLabel();
            this.tbDataSourceSelect = new System.Windows.Forms.TabPage();
            this.rtbReselectWarning = new System.Windows.Forms.RichTextBox();
            this.gbSpreadsheet = new System.Windows.Forms.GroupBox();
            this.cbSpreadsheetTable = new System.Windows.Forms.ComboBox();
            this.lTable = new System.Windows.Forms.Label();
            this.bSelectSpreadsheet = new System.Windows.Forms.Button();
            this.tbSpreadsheet = new System.Windows.Forms.TextBox();
            this.gbSpreadsheetInformation = new System.Windows.Forms.GroupBox();
            this.lPossibleColumnsValue = new System.Windows.Forms.Label();
            this.lPossibleColumns = new System.Windows.Forms.Label();
            this.lColumns = new System.Windows.Forms.Label();
            this.cbDataSourceColumnsExampleSpreadsheet = new System.Windows.Forms.ComboBox();
            this.lPossibleRowsValue = new System.Windows.Forms.Label();
            this.lPossibleRows = new System.Windows.Forms.Label();
            this.tpPDFSelect = new System.Windows.Forms.TabPage();
            this.tbFormTyp = new System.Windows.Forms.TextBox();
            this.dgvBulkPDF = new System.Windows.Forms.DataGridView();
            this.ColField = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColTyp = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColOption = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tbPDF = new System.Windows.Forms.TextBox();
            this.bSelectPDF = new System.Windows.Forms.Button();
            this.tpFinish = new System.Windows.Forms.TabPage();
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.cFilterCheck = new System.Windows.Forms.ComboBox();
            this.cFilterField = new System.Windows.Forms.ComboBox();
            this.cFilter = new System.Windows.Forms.CheckBox();
            this.cFlatten = new System.Windows.Forms.CheckBox();
            this.cFusion = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cRowsPerPag = new System.Windows.Forms.NumericUpDown();
            this.cGrouped = new System.Windows.Forms.CheckBox();
            this.cbUnicode = new System.Windows.Forms.CheckBox();
            this.cbFinalize = new System.Windows.Forms.CheckBox();
            this.cGroupByColumn = new System.Windows.Forms.ComboBox();
            this.gbOuput = new System.Windows.Forms.GroupBox();
            this.tbOutputDir = new System.Windows.Forms.TextBox();
            this.bSelectOutputPath = new System.Windows.Forms.Button();
            this.bSaveConfiguration = new System.Windows.Forms.Button();
            this.gbFilename = new System.Windows.Forms.GroupBox();
            this.cUsePag = new System.Windows.Forms.CheckBox();
            this.cSufix2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cUseGroup = new System.Windows.Forms.CheckBox();
            this.cUseRowNumber = new System.Windows.Forms.CheckBox();
            this.lPlus3 = new System.Windows.Forms.Label();
            this.tbExampleFilename = new System.Windows.Forms.TextBox();
            this.lExampleFilename = new System.Windows.Forms.Label();
            this.cUseValueFromDataSource = new System.Windows.Forms.CheckBox();
            this.lPlus2 = new System.Windows.Forms.Label();
            this.lPlus1 = new System.Windows.Forms.Label();
            this.cSufix1 = new System.Windows.Forms.TextBox();
            this.cDataSourceColumnsFilename = new System.Windows.Forms.ComboBox();
            this.cPrefix = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.wizardPages.SuspendLayout();
            this.tpWelcome.SuspendLayout();
            this.tbDataSourceSelect.SuspendLayout();
            this.gbSpreadsheet.SuspendLayout();
            this.gbSpreadsheetInformation.SuspendLayout();
            this.tpPDFSelect.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBulkPDF)).BeginInit();
            this.tpFinish.SuspendLayout();
            this.gbOptions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.cRowsPerPag)).BeginInit();
            this.gbOuput.SuspendLayout();
            this.gbFilename.SuspendLayout();
            this.SuspendLayout();
            // 
            // bNext
            // 
            resources.ApplyResources(this.bNext, "bNext");
            this.bNext.Name = "bNext";
            this.bNext.UseVisualStyleBackColor = true;
            this.bNext.Click += new System.EventHandler(this.bNext_Click);
            // 
            // bBack
            // 
            resources.ApplyResources(this.bBack, "bBack");
            this.bBack.Name = "bBack";
            this.bBack.UseVisualStyleBackColor = true;
            this.bBack.Click += new System.EventHandler(this.bBack_Click);
            // 
            // bFinish
            // 
            resources.ApplyResources(this.bFinish, "bFinish");
            this.bFinish.Name = "bFinish";
            this.bFinish.UseVisualStyleBackColor = true;
            this.bFinish.Click += new System.EventHandler(this.bFinish_Click);
            // 
            // pictureBox1
            // 
            resources.ApplyResources(this.pictureBox1, "pictureBox1");
            this.pictureBox1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.TabStop = false;
            // 
            // bDonate
            // 
            resources.ApplyResources(this.bDonate, "bDonate");
            this.bDonate.Name = "bDonate";
            this.bDonate.UseVisualStyleBackColor = true;
            this.bDonate.Click += new System.EventHandler(this.bDonate_Click);
            // 
            // seperatorLine1
            // 
            resources.ApplyResources(this.seperatorLine1, "seperatorLine1");
            this.seperatorLine1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.seperatorLine1.Name = "seperatorLine1";
            // 
            // wizardPages
            // 
            resources.ApplyResources(this.wizardPages, "wizardPages");
            this.wizardPages.Controls.Add(this.tpWelcome);
            this.wizardPages.Controls.Add(this.tbDataSourceSelect);
            this.wizardPages.Controls.Add(this.tpPDFSelect);
            this.wizardPages.Controls.Add(this.tpFinish);
            this.wizardPages.Name = "wizardPages";
            this.wizardPages.SelectedIndex = 0;
            this.wizardPages.TabStop = false;
            this.wizardPages.SelectedIndexChanged += new System.EventHandler(this.wizardPages_SelectedIndexChanged);
            // 
            // tpWelcome
            // 
            resources.ApplyResources(this.tpWelcome, "tpWelcome");
            this.tpWelcome.BackColor = System.Drawing.SystemColors.Control;
            this.tpWelcome.Controls.Add(this.llSupport);
            this.tpWelcome.Controls.Add(this.llBulkPDFde);
            this.tpWelcome.Controls.Add(this.llLicenses);
            this.tpWelcome.Controls.Add(this.rtbWelcome);
            this.tpWelcome.Controls.Add(this.lVersion);
            this.tpWelcome.Controls.Add(this.lBulkPDF);
            this.tpWelcome.Controls.Add(this.bLoadConfiguration);
            this.tpWelcome.Controls.Add(this.llDokumentation);
            this.tpWelcome.Name = "tpWelcome";
            // 
            // llSupport
            // 
            resources.ApplyResources(this.llSupport, "llSupport");
            this.llSupport.Name = "llSupport";
            this.llSupport.TabStop = true;
            this.llSupport.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llSupport_LinkClicked);
            // 
            // llBulkPDFde
            // 
            resources.ApplyResources(this.llBulkPDFde, "llBulkPDFde");
            this.llBulkPDFde.Name = "llBulkPDFde";
            this.llBulkPDFde.TabStop = true;
            this.llBulkPDFde.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llBulkPDFde_LinkClicked);
            // 
            // llLicenses
            // 
            resources.ApplyResources(this.llLicenses, "llLicenses");
            this.llLicenses.BackColor = System.Drawing.Color.Transparent;
            this.llLicenses.Name = "llLicenses";
            this.llLicenses.TabStop = true;
            this.llLicenses.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llLicenses_LinkClicked);
            // 
            // rtbWelcome
            // 
            resources.ApplyResources(this.rtbWelcome, "rtbWelcome");
            this.rtbWelcome.BackColor = System.Drawing.SystemColors.Control;
            this.rtbWelcome.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtbWelcome.Name = "rtbWelcome";
            this.rtbWelcome.ReadOnly = true;
            // 
            // lVersion
            // 
            resources.ApplyResources(this.lVersion, "lVersion");
            this.lVersion.BackColor = System.Drawing.Color.Transparent;
            this.lVersion.Name = "lVersion";
            // 
            // lBulkPDF
            // 
            resources.ApplyResources(this.lBulkPDF, "lBulkPDF");
            this.lBulkPDF.Name = "lBulkPDF";
            // 
            // bLoadConfiguration
            // 
            resources.ApplyResources(this.bLoadConfiguration, "bLoadConfiguration");
            this.bLoadConfiguration.Name = "bLoadConfiguration";
            this.bLoadConfiguration.UseVisualStyleBackColor = true;
            this.bLoadConfiguration.Click += new System.EventHandler(this.bLoadConfiguration_Click);
            // 
            // llDokumentation
            // 
            resources.ApplyResources(this.llDokumentation, "llDokumentation");
            this.llDokumentation.Name = "llDokumentation";
            this.llDokumentation.TabStop = true;
            this.llDokumentation.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llDokumentation_LinkClicked);
            // 
            // tbDataSourceSelect
            // 
            resources.ApplyResources(this.tbDataSourceSelect, "tbDataSourceSelect");
            this.tbDataSourceSelect.BackColor = System.Drawing.SystemColors.Control;
            this.tbDataSourceSelect.Controls.Add(this.rtbReselectWarning);
            this.tbDataSourceSelect.Controls.Add(this.gbSpreadsheet);
            this.tbDataSourceSelect.Controls.Add(this.gbSpreadsheetInformation);
            this.tbDataSourceSelect.Name = "tbDataSourceSelect";
            // 
            // rtbReselectWarning
            // 
            resources.ApplyResources(this.rtbReselectWarning, "rtbReselectWarning");
            this.rtbReselectWarning.BackColor = System.Drawing.SystemColors.Control;
            this.rtbReselectWarning.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtbReselectWarning.Name = "rtbReselectWarning";
            // 
            // gbSpreadsheet
            // 
            resources.ApplyResources(this.gbSpreadsheet, "gbSpreadsheet");
            this.gbSpreadsheet.Controls.Add(this.cbSpreadsheetTable);
            this.gbSpreadsheet.Controls.Add(this.lTable);
            this.gbSpreadsheet.Controls.Add(this.bSelectSpreadsheet);
            this.gbSpreadsheet.Controls.Add(this.tbSpreadsheet);
            this.gbSpreadsheet.Name = "gbSpreadsheet";
            this.gbSpreadsheet.TabStop = false;
            // 
            // cbSpreadsheetTable
            // 
            resources.ApplyResources(this.cbSpreadsheetTable, "cbSpreadsheetTable");
            this.cbSpreadsheetTable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSpreadsheetTable.FormattingEnabled = true;
            this.cbSpreadsheetTable.Name = "cbSpreadsheetTable";
            this.cbSpreadsheetTable.SelectedIndexChanged += new System.EventHandler(this.cbTable_SelectedIndexChanged);
            // 
            // lTable
            // 
            resources.ApplyResources(this.lTable, "lTable");
            this.lTable.Name = "lTable";
            // 
            // bSelectSpreadsheet
            // 
            resources.ApplyResources(this.bSelectSpreadsheet, "bSelectSpreadsheet");
            this.bSelectSpreadsheet.Name = "bSelectSpreadsheet";
            this.bSelectSpreadsheet.UseVisualStyleBackColor = true;
            this.bSelectSpreadsheet.Click += new System.EventHandler(this.bSelectSpreadsheet_Click);
            // 
            // tbSpreadsheet
            // 
            resources.ApplyResources(this.tbSpreadsheet, "tbSpreadsheet");
            this.tbSpreadsheet.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tbSpreadsheet.Name = "tbSpreadsheet";
            // 
            // gbSpreadsheetInformation
            // 
            resources.ApplyResources(this.gbSpreadsheetInformation, "gbSpreadsheetInformation");
            this.gbSpreadsheetInformation.Controls.Add(this.lPossibleColumnsValue);
            this.gbSpreadsheetInformation.Controls.Add(this.lPossibleColumns);
            this.gbSpreadsheetInformation.Controls.Add(this.lColumns);
            this.gbSpreadsheetInformation.Controls.Add(this.cbDataSourceColumnsExampleSpreadsheet);
            this.gbSpreadsheetInformation.Controls.Add(this.lPossibleRowsValue);
            this.gbSpreadsheetInformation.Controls.Add(this.lPossibleRows);
            this.gbSpreadsheetInformation.Name = "gbSpreadsheetInformation";
            this.gbSpreadsheetInformation.TabStop = false;
            // 
            // lPossibleColumnsValue
            // 
            resources.ApplyResources(this.lPossibleColumnsValue, "lPossibleColumnsValue");
            this.lPossibleColumnsValue.Name = "lPossibleColumnsValue";
            // 
            // lPossibleColumns
            // 
            resources.ApplyResources(this.lPossibleColumns, "lPossibleColumns");
            this.lPossibleColumns.Name = "lPossibleColumns";
            // 
            // lColumns
            // 
            resources.ApplyResources(this.lColumns, "lColumns");
            this.lColumns.Name = "lColumns";
            // 
            // cbDataSourceColumnsExampleSpreadsheet
            // 
            resources.ApplyResources(this.cbDataSourceColumnsExampleSpreadsheet, "cbDataSourceColumnsExampleSpreadsheet");
            this.cbDataSourceColumnsExampleSpreadsheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbDataSourceColumnsExampleSpreadsheet.FormattingEnabled = true;
            this.cbDataSourceColumnsExampleSpreadsheet.Name = "cbDataSourceColumnsExampleSpreadsheet";
            // 
            // lPossibleRowsValue
            // 
            resources.ApplyResources(this.lPossibleRowsValue, "lPossibleRowsValue");
            this.lPossibleRowsValue.Name = "lPossibleRowsValue";
            // 
            // lPossibleRows
            // 
            resources.ApplyResources(this.lPossibleRows, "lPossibleRows");
            this.lPossibleRows.Name = "lPossibleRows";
            // 
            // tpPDFSelect
            // 
            resources.ApplyResources(this.tpPDFSelect, "tpPDFSelect");
            this.tpPDFSelect.BackColor = System.Drawing.SystemColors.Control;
            this.tpPDFSelect.Controls.Add(this.tbFormTyp);
            this.tpPDFSelect.Controls.Add(this.dgvBulkPDF);
            this.tpPDFSelect.Controls.Add(this.tbPDF);
            this.tpPDFSelect.Controls.Add(this.bSelectPDF);
            this.tpPDFSelect.Name = "tpPDFSelect";
            // 
            // tbFormTyp
            // 
            resources.ApplyResources(this.tbFormTyp, "tbFormTyp");
            this.tbFormTyp.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tbFormTyp.Name = "tbFormTyp";
            // 
            // dgvBulkPDF
            // 
            resources.ApplyResources(this.dgvBulkPDF, "dgvBulkPDF");
            this.dgvBulkPDF.AllowUserToAddRows = false;
            this.dgvBulkPDF.AllowUserToDeleteRows = false;
            this.dgvBulkPDF.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvBulkPDF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvBulkPDF.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColField,
            this.ColTyp,
            this.ColValue,
            this.ColOption});
            this.dgvBulkPDF.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dgvBulkPDF.Name = "dgvBulkPDF";
            this.dgvBulkPDF.ReadOnly = true;
            this.dgvBulkPDF.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBulkPDF_CellContentClick);
            // 
            // ColField
            // 
            this.ColField.FillWeight = 67.96245F;
            resources.ApplyResources(this.ColField, "ColField");
            this.ColField.Name = "ColField";
            this.ColField.ReadOnly = true;
            // 
            // ColTyp
            // 
            this.ColTyp.FillWeight = 109.3781F;
            resources.ApplyResources(this.ColTyp, "ColTyp");
            this.ColTyp.Name = "ColTyp";
            this.ColTyp.ReadOnly = true;
            // 
            // ColValue
            // 
            this.ColValue.FillWeight = 108.8228F;
            resources.ApplyResources(this.ColValue, "ColValue");
            this.ColValue.Name = "ColValue";
            this.ColValue.ReadOnly = true;
            // 
            // ColOption
            // 
            this.ColOption.FillWeight = 109.3781F;
            resources.ApplyResources(this.ColOption, "ColOption");
            this.ColOption.Name = "ColOption";
            this.ColOption.ReadOnly = true;
            this.ColOption.Text = "";
            // 
            // tbPDF
            // 
            resources.ApplyResources(this.tbPDF, "tbPDF");
            this.tbPDF.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tbPDF.Name = "tbPDF";
            // 
            // bSelectPDF
            // 
            resources.ApplyResources(this.bSelectPDF, "bSelectPDF");
            this.bSelectPDF.Name = "bSelectPDF";
            this.bSelectPDF.UseVisualStyleBackColor = true;
            this.bSelectPDF.Click += new System.EventHandler(this.bSelectPDF_Click);
            // 
            // tpFinish
            // 
            resources.ApplyResources(this.tpFinish, "tpFinish");
            this.tpFinish.BackColor = System.Drawing.SystemColors.Control;
            this.tpFinish.Controls.Add(this.gbOptions);
            this.tpFinish.Controls.Add(this.gbOuput);
            this.tpFinish.Controls.Add(this.bSaveConfiguration);
            this.tpFinish.Controls.Add(this.gbFilename);
            this.tpFinish.Name = "tpFinish";
            // 
            // gbOptions
            // 
            resources.ApplyResources(this.gbOptions, "gbOptions");
            this.gbOptions.Controls.Add(this.cFilterCheck);
            this.gbOptions.Controls.Add(this.cFilterField);
            this.gbOptions.Controls.Add(this.cFilter);
            this.gbOptions.Controls.Add(this.cFlatten);
            this.gbOptions.Controls.Add(this.cFusion);
            this.gbOptions.Controls.Add(this.label4);
            this.gbOptions.Controls.Add(this.cRowsPerPag);
            this.gbOptions.Controls.Add(this.cGrouped);
            this.gbOptions.Controls.Add(this.cbUnicode);
            this.gbOptions.Controls.Add(this.cbFinalize);
            this.gbOptions.Controls.Add(this.cGroupByColumn);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.TabStop = false;
            // 
            // cFilterCheck
            // 
            resources.ApplyResources(this.cFilterCheck, "cFilterCheck");
            this.cFilterCheck.FormattingEnabled = true;
            this.cFilterCheck.Name = "cFilterCheck";
            // 
            // cFilterField
            // 
            resources.ApplyResources(this.cFilterField, "cFilterField");
            this.cFilterField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cFilterField.FormattingEnabled = true;
            this.cFilterField.Name = "cFilterField";
            // 
            // cFilter
            // 
            resources.ApplyResources(this.cFilter, "cFilter");
            this.cFilter.Name = "cFilter";
            this.cFilter.UseVisualStyleBackColor = true;
            this.cFilter.CheckedChanged += new System.EventHandler(this.cFilter_CheckedChanged);
            // 
            // cFlatten
            // 
            resources.ApplyResources(this.cFlatten, "cFlatten");
            this.cFlatten.Name = "cFlatten";
            this.cFlatten.UseVisualStyleBackColor = true;
            // 
            // cFusion
            // 
            resources.ApplyResources(this.cFusion, "cFusion");
            this.cFusion.Name = "cFusion";
            this.cFusion.UseVisualStyleBackColor = true;
            this.cFusion.CheckedChanged += new System.EventHandler(this.cCheckOptionsChanged);
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // cRowsPerPag
            // 
            resources.ApplyResources(this.cRowsPerPag, "cRowsPerPag");
            this.cRowsPerPag.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.cRowsPerPag.Name = "cRowsPerPag";
            this.cRowsPerPag.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // cGrouped
            // 
            resources.ApplyResources(this.cGrouped, "cGrouped");
            this.cGrouped.Name = "cGrouped";
            this.cGrouped.UseVisualStyleBackColor = true;
            this.cGrouped.CheckedChanged += new System.EventHandler(this.cCheckOptionsChanged);
            // 
            // cbUnicode
            // 
            resources.ApplyResources(this.cbUnicode, "cbUnicode");
            this.cbUnicode.Name = "cbUnicode";
            this.cbUnicode.UseVisualStyleBackColor = true;
            // 
            // cbFinalize
            // 
            resources.ApplyResources(this.cbFinalize, "cbFinalize");
            this.cbFinalize.Name = "cbFinalize";
            this.cbFinalize.UseVisualStyleBackColor = true;
            // 
            // cGroupByColumn
            // 
            resources.ApplyResources(this.cGroupByColumn, "cGroupByColumn");
            this.cGroupByColumn.FormattingEnabled = true;
            this.cGroupByColumn.Name = "cGroupByColumn";
            this.cGroupByColumn.SelectedIndexChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // gbOuput
            // 
            resources.ApplyResources(this.gbOuput, "gbOuput");
            this.gbOuput.Controls.Add(this.tbOutputDir);
            this.gbOuput.Controls.Add(this.bSelectOutputPath);
            this.gbOuput.Name = "gbOuput";
            this.gbOuput.TabStop = false;
            // 
            // tbOutputDir
            // 
            resources.ApplyResources(this.tbOutputDir, "tbOutputDir");
            this.tbOutputDir.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tbOutputDir.Name = "tbOutputDir";
            // 
            // bSelectOutputPath
            // 
            resources.ApplyResources(this.bSelectOutputPath, "bSelectOutputPath");
            this.bSelectOutputPath.Name = "bSelectOutputPath";
            this.bSelectOutputPath.UseVisualStyleBackColor = true;
            this.bSelectOutputPath.Click += new System.EventHandler(this.bSelectOutputPath_Click);
            // 
            // bSaveConfiguration
            // 
            resources.ApplyResources(this.bSaveConfiguration, "bSaveConfiguration");
            this.bSaveConfiguration.Name = "bSaveConfiguration";
            this.bSaveConfiguration.UseVisualStyleBackColor = true;
            this.bSaveConfiguration.Click += new System.EventHandler(this.bSaveConfiguration_Click);
            // 
            // gbFilename
            // 
            resources.ApplyResources(this.gbFilename, "gbFilename");
            this.gbFilename.Controls.Add(this.cUsePag);
            this.gbFilename.Controls.Add(this.cSufix2);
            this.gbFilename.Controls.Add(this.label3);
            this.gbFilename.Controls.Add(this.label2);
            this.gbFilename.Controls.Add(this.label1);
            this.gbFilename.Controls.Add(this.cUseGroup);
            this.gbFilename.Controls.Add(this.cUseRowNumber);
            this.gbFilename.Controls.Add(this.lPlus3);
            this.gbFilename.Controls.Add(this.tbExampleFilename);
            this.gbFilename.Controls.Add(this.lExampleFilename);
            this.gbFilename.Controls.Add(this.cUseValueFromDataSource);
            this.gbFilename.Controls.Add(this.lPlus2);
            this.gbFilename.Controls.Add(this.lPlus1);
            this.gbFilename.Controls.Add(this.cSufix1);
            this.gbFilename.Controls.Add(this.cDataSourceColumnsFilename);
            this.gbFilename.Controls.Add(this.cPrefix);
            this.gbFilename.Name = "gbFilename";
            this.gbFilename.TabStop = false;
            // 
            // cUsePag
            // 
            resources.ApplyResources(this.cUsePag, "cUsePag");
            this.cUsePag.Name = "cUsePag";
            this.cUsePag.UseVisualStyleBackColor = true;
            this.cUsePag.CheckedChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // cSufix2
            // 
            resources.ApplyResources(this.cSufix2, "cSufix2");
            this.cSufix2.Name = "cSufix2";
            this.cSufix2.TextChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // cUseGroup
            // 
            resources.ApplyResources(this.cUseGroup, "cUseGroup");
            this.cUseGroup.Name = "cUseGroup";
            this.cUseGroup.UseVisualStyleBackColor = true;
            this.cUseGroup.CheckedChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // cUseRowNumber
            // 
            resources.ApplyResources(this.cUseRowNumber, "cUseRowNumber");
            this.cUseRowNumber.Name = "cUseRowNumber";
            this.cUseRowNumber.UseVisualStyleBackColor = true;
            this.cUseRowNumber.CheckedChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // lPlus3
            // 
            resources.ApplyResources(this.lPlus3, "lPlus3");
            this.lPlus3.Name = "lPlus3";
            // 
            // tbExampleFilename
            // 
            resources.ApplyResources(this.tbExampleFilename, "tbExampleFilename");
            this.tbExampleFilename.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tbExampleFilename.Name = "tbExampleFilename";
            // 
            // lExampleFilename
            // 
            resources.ApplyResources(this.lExampleFilename, "lExampleFilename");
            this.lExampleFilename.Name = "lExampleFilename";
            // 
            // cUseValueFromDataSource
            // 
            resources.ApplyResources(this.cUseValueFromDataSource, "cUseValueFromDataSource");
            this.cUseValueFromDataSource.Name = "cUseValueFromDataSource";
            this.cUseValueFromDataSource.UseVisualStyleBackColor = true;
            this.cUseValueFromDataSource.CheckedChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // lPlus2
            // 
            resources.ApplyResources(this.lPlus2, "lPlus2");
            this.lPlus2.Name = "lPlus2";
            // 
            // lPlus1
            // 
            resources.ApplyResources(this.lPlus1, "lPlus1");
            this.lPlus1.Name = "lPlus1";
            // 
            // cSufix1
            // 
            resources.ApplyResources(this.cSufix1, "cSufix1");
            this.cSufix1.Name = "cSufix1";
            this.cSufix1.TextChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // cDataSourceColumnsFilename
            // 
            resources.ApplyResources(this.cDataSourceColumnsFilename, "cDataSourceColumnsFilename");
            this.cDataSourceColumnsFilename.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cDataSourceColumnsFilename.FormattingEnabled = true;
            this.cDataSourceColumnsFilename.Name = "cDataSourceColumnsFilename";
            this.cDataSourceColumnsFilename.SelectedIndexChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // cPrefix
            // 
            resources.ApplyResources(this.cPrefix, "cPrefix");
            this.cPrefix.Name = "cPrefix";
            this.cPrefix.TextChanged += new System.EventHandler(this.cFilenameChanged);
            // 
            // MainForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.bDonate);
            this.Controls.Add(this.seperatorLine1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.bFinish);
            this.Controls.Add(this.bBack);
            this.Controls.Add(this.bNext);
            this.Controls.Add(this.wizardPages);
            this.Name = "MainForm";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.wizardPages.ResumeLayout(false);
            this.tpWelcome.ResumeLayout(false);
            this.tpWelcome.PerformLayout();
            this.tbDataSourceSelect.ResumeLayout(false);
            this.gbSpreadsheet.ResumeLayout(false);
            this.gbSpreadsheet.PerformLayout();
            this.gbSpreadsheetInformation.ResumeLayout(false);
            this.gbSpreadsheetInformation.PerformLayout();
            this.tpPDFSelect.ResumeLayout(false);
            this.tpPDFSelect.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBulkPDF)).EndInit();
            this.tpFinish.ResumeLayout(false);
            this.gbOptions.ResumeLayout(false);
            this.gbOptions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.cRowsPerPag)).EndInit();
            this.gbOuput.ResumeLayout(false);
            this.gbOuput.PerformLayout();
            this.gbFilename.ResumeLayout(false);
            this.gbFilename.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private WizardPages wizardPages;
        private System.Windows.Forms.TabPage tbDataSourceSelect;
        private System.Windows.Forms.Button bNext;
        private System.Windows.Forms.Button bBack;
        private System.Windows.Forms.TabPage tpPDFSelect;
        private SeperatorLine seperatorLine1;
        private System.Windows.Forms.Button bFinish;
        private System.Windows.Forms.GroupBox gbSpreadsheetInformation;
        private System.Windows.Forms.Label lPossibleRowsValue;
        private System.Windows.Forms.Label lPossibleRows;
        private System.Windows.Forms.Button bSelectPDF;
        private System.Windows.Forms.TabPage tpFinish;
        private System.Windows.Forms.TextBox tbPDF;
        private System.Windows.Forms.DataGridView dgvBulkPDF;
        private System.Windows.Forms.TextBox tbSpreadsheet;
        private System.Windows.Forms.Button bSelectSpreadsheet;
        private System.Windows.Forms.GroupBox gbFilename;
        private System.Windows.Forms.Button bSaveConfiguration;
        private System.Windows.Forms.TextBox cSufix1;
        private System.Windows.Forms.ComboBox cDataSourceColumnsFilename;
        private System.Windows.Forms.TextBox cPrefix;
        private System.Windows.Forms.Label lPlus2;
        private System.Windows.Forms.Label lPlus1;
        private System.Windows.Forms.GroupBox gbOuput;
        private System.Windows.Forms.TextBox tbOutputDir;
        private System.Windows.Forms.Button bSelectOutputPath;
        private System.Windows.Forms.CheckBox cUseValueFromDataSource;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox tbExampleFilename;
        private System.Windows.Forms.Label lExampleFilename;
        private System.Windows.Forms.Label lPlus3;
        private System.Windows.Forms.CheckBox cUseRowNumber;
        private System.Windows.Forms.Label lColumns;
        private System.Windows.Forms.ComboBox cbDataSourceColumnsExampleSpreadsheet;
        private System.Windows.Forms.GroupBox gbSpreadsheet;
        private System.Windows.Forms.Label lTable;
        private System.Windows.Forms.ComboBox cbSpreadsheetTable;
        private System.Windows.Forms.Button bDonate;
        private System.Windows.Forms.Label lPossibleColumnsValue;
        private System.Windows.Forms.Label lPossibleColumns;
        private System.Windows.Forms.GroupBox gbOptions;
        private System.Windows.Forms.CheckBox cbFinalize;
        private System.Windows.Forms.TextBox tbFormTyp;
        private System.Windows.Forms.RichTextBox rtbReselectWarning;
        private System.Windows.Forms.CheckBox cbUnicode;
        private System.Windows.Forms.TabPage tpWelcome;
        private System.Windows.Forms.LinkLabel llSupport;
        private System.Windows.Forms.LinkLabel llBulkPDFde;
        private System.Windows.Forms.LinkLabel llLicenses;
        private System.Windows.Forms.RichTextBox rtbWelcome;
        private System.Windows.Forms.Label lVersion;
        private System.Windows.Forms.Label lBulkPDF;
        private System.Windows.Forms.Button bLoadConfiguration;
        private System.Windows.Forms.LinkLabel llDokumentation;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cGroupByColumn;
        private System.Windows.Forms.CheckBox cUseGroup;
        private System.Windows.Forms.TextBox cSufix2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox cUsePag;
        private System.Windows.Forms.CheckBox cGrouped;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown cRowsPerPag;
        private System.Windows.Forms.CheckBox cFusion;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColField;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColTyp;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColValue;
        private System.Windows.Forms.DataGridViewButtonColumn ColOption;
        private System.Windows.Forms.ComboBox cFilterCheck;
        private System.Windows.Forms.ComboBox cFilterField;
        private System.Windows.Forms.CheckBox cFilter;
        private System.Windows.Forms.CheckBox cFlatten;
    }
}