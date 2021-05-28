using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace BulkPDF
{
    public partial class MainForm : Form
    {
        PDF pdf;
        IDataSource dataSource;
        Dictionary<string, PDFField> pdfFields = new Dictionary<string, PDFField>();
        ProgressForm progressForm;
        Opciones opt;
        //int tempSelectedIndex;
        
        
        public MainForm()
        {
            InitializeComponent();
            this.MinimumSize = new Size(500, 400);

            lVersion.Text = Application.ProductVersion.ToString();
        }

        /**************************************************/
        #region WizardPage
        /**************************************************/

        private void MainForm_Load(object sender, EventArgs e)
        {
            wizardPages.SelectedIndex = 1;
            wizardPages.SelectedIndex = 0;
            PopulateFilterOperators();
        }


        private void bNext_Click(object sender, EventArgs e)
        {
            if (IsNextPageOk())
            {
                this.SuspendLayout();
                if (wizardPages.SelectedIndex < wizardPages.TabPages.Count)
                    wizardPages.SelectedIndex += 1;
                this.ResumeLayout();
            }
        }

        private void bBack_Click(object sender, EventArgs e)
        {
            this.SuspendLayout();
            if (wizardPages.SelectedIndex > 0)
                wizardPages.SelectedIndex -= 1;
            this.ResumeLayout();
        }

        private void wizardPages_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wizardPages.SelectedIndex == 0)
            {
                bBack.Hide();
            }
            else
            {
                bBack.Show();
            }

            if (wizardPages.SelectedIndex != (wizardPages.TabPages.Count - 1))
            {
                bFinish.Hide();
                bNext.Show();
            }
            else
            {
                bNext.Show();
                bFinish.Show();
            }
        }

        /**************************************************/
        #endregion
        /**************************************************/

        private bool IsNextPageOk()
        {
            switch (wizardPages.SelectedTab.Text)
            {
                case "SpreadsheetSelect":
                    if (dataSource == null)
                    {
                        MessageBox.Show(Properties.Resources.MessageSelectSpreadsheet);
                        return false;
                    }
                    if (dataSource.PossibleRows == 0)
                    {
                        MessageBox.Show(Properties.Resources.MessageNoUsableRows);
                        return false;
                    }
                    break;
                case "PDFSelect":
                    if (pdf == null)
                    {
                        MessageBox.Show(Properties.Resources.MessageNoPDFSelected);
                        return false;
                    }
                    break;
            }

            return true;
        }


        private void bDonate_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://bulkpdf.de/donate");
        }

        private void llDokumentation_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://bulkpdf.de/documentation");
        }

        private void llLicenses_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var licenses = new Licenses();
            licenses.ShowDialog();
        }

        private void llBulkPDFde_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://bulkpdf.de/");
        }

        /**************************************************/
        #region Fill
        /**************************************************/

        private void bFinish_Click(object sender, EventArgs e)
        {
            SincronizaOpciones();
            if (!String.IsNullOrEmpty(tbOutputDir.Text))
            {
                if (Directory.Exists(tbOutputDir.Text))
                {
                    PDFFiller filler = new PDFFiller(pdf, dataSource, opt, ConcatFilename);
                    int nArchivos = opt.Fusion ? 1 : filler.IsFilenameUnique();// IsFilenameUnique();
                    if (nArchivos>0)
                    {
                        DialogResult dialogResult = MessageBox.Show(String.Format(Properties.Resources.MessageCreateNPDFFilesInDir, nArchivos, tbOutputDir.Text), Properties.Resources.MessageAreYouSure, MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            this.SuspendLayout();

                            progressForm = new ProgressForm();
                            progressForm.Show();
                            BackgroundWorker backGroundWorker = new BackgroundWorker();
                            backGroundWorker.DoWork += backGroundWorker_DoWork;
                            backGroundWorker.RunWorkerAsync(filler);

                            this.ResumeLayout();
                        }
                    }
                    else
                    {
                        MessageBox.Show(Properties.Resources.MessageFilenameNotUnique);
                    }
                }
                else
                {
                    MessageBox.Show(Properties.Resources.MessageOutputDirNotExist);
                }
            }
            else
            {
                MessageBox.Show(Properties.Resources.MessageNoOutputDirSelected);
            }
        }

        void backGroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                PDFFiller filler = (PDFFiller)e.Argument;
                if(opt.Fusion)
                    pdf.CreateOneFile(filler, pdfFields, progressForm.SetPercent, progressForm.GetIsAborted);
                else
                    pdf.CreateFiles(filler, pdfFields, progressForm.SetPercent, progressForm.GetIsAborted);
                // OLD --CreateFiles(pdf, dataSource, pdfFields, opt, ConcatFilename, progressForm.SetPercent, progressForm.GetIsAborted);
            }
            catch (Exception ex)
            {
                ExceptionHandler.Throw(Properties.Resources.ExceptionPDFFileAlreadyExistsAndInUse, ex.ToString());
                
                this.Invoke((MethodInvoker)delegate {
                    progressForm.Close();
                });
            }
        }

        private string ConcatFilename(int pageGroup)
        {
            return Opciones.ConcatFilename(dataSource, opt, pageGroup);
        }

/*
        private int IsFilenameUnique()
        {
            // Is dataSource unique?
            //if (cUseValueFromDataSource.Checked || cUseGroup.Checked)
            //{
                return PDFFiller.VerificaFiles(dataSource, opt, ConcatFilename);
            //}

            //return 0;
        }
*/
        private void bSelectOutputPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                tbOutputDir.Text = folderBrowserDialog.SelectedPath;
                opt.OutputDir = tbOutputDir.Text;
            }
        }

        /**************************************************/
        #endregion
        /**************************************************/

        /**************************************************/
        #region PDFSelect
        /**************************************************/

        private void bSelectPDF_Click(object sender, EventArgs e)
        {
            // Select File
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "PDF|*.pdf";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // PDF
                if (pdf != null)
                    pdf.Close();
                OpenPDF(openFileDialog.FileName);
            }
        }

        private bool OpenPDF(string pdfPath)
        {
            try
            {
                ResetPDF();

                pdf = PDF.Open(pdfPath);

                // Fill DataGridView
                tbPDF.Text = pdfPath;
                tbFormTyp.Text = pdf.getTipoTexto();
                foreach (PDFField pdfField in pdf.ListFields())
                {
                    dgvBulkPDF.Rows.Add();
                    int row = dgvBulkPDF.Rows.Count - 1;

                    dgvBulkPDF.Rows[row].Cells["ColField"].Value = pdfField.Name;
                    dgvBulkPDF.Rows[row].Cells["ColTyp"].Value = pdfField.Typ;
                    dgvBulkPDF.Rows[row].Cells["ColValue"].Value = pdfField.CurrentValue;
                    pdfFields.Add(pdfField.Name, pdfField);
                    
                    var dgvButtonCell = new DataGridViewButtonCell();
                    dgvButtonCell.Value = Properties.Resources.CellButtonSelect;
                    dgvBulkPDF.Rows[row].Cells["ColOption"] = dgvButtonCell;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.Throw(String.Format(Properties.Resources.ExceptionPDFIsCorrupted, pdfPath), ex.ToString());
                ResetPDF();
            }
            return false;
        }

        private void ResetPDF()
        {
            pdf = null;
            dgvBulkPDF.Rows.Clear();
            pdfFields.Clear();
            tbPDF.Text = "";
            tbFormTyp.Text = "";
        }

        private void dgvBulkPDF_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
            {
                var fieldOptionForm = new FieldOptionForm(new Point(this.Location.X + Convert.ToInt32(this.Width / 2.5), this.Location.Y + Convert.ToInt32(this.Height / 2.5))
                    , pdfFields[(string)dgvBulkPDF.Rows[e.RowIndex].Cells["ColField"].Value], dataSource.Columns);
                fieldOptionForm.ShowDialog();
                if (fieldOptionForm.ShouldBeSaved)
                {
                    pdfFields[fieldOptionForm.PDFField.Name] = fieldOptionForm.PDFField;
                    if (fieldOptionForm.PDFField.UseValueFromDataSource)
                    {
                        string value = pdfFields[fieldOptionForm.PDFField.Name].DataSourceValue;
                        if (pdfFields[fieldOptionForm.PDFField.Name].MakeReadOnly)
                            value = "[#]" + value;

                        dgvBulkPDF.Rows[e.RowIndex].Cells["ColValue"].Value = value;
                    }
                    else
                    {
                        dgvBulkPDF.Rows[e.RowIndex].Cells["ColValue"].Value = pdfFields[fieldOptionForm.PDFField.Name].CurrentValue;
                    }
                }
            }
        }

        /**************************************************/
        #endregion
        /**************************************************/

        /**************************************************/
        #region Save & Load
        /**************************************************/

        private void bSaveConfiguration_Click(object sender, EventArgs e)
        {
            SincronizaOpciones();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "BulkPDF Options|*.bulkpdf";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var xmlWriterSettings = new XmlWriterSettings();
                xmlWriterSettings.Indent = true;
                xmlWriterSettings.IndentChars = "  ";
                var xmlWriter = XmlWriter.Create(saveFileDialog.FileName, xmlWriterSettings);
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement("BulkPDF"); // <BulkPDF>
                xmlWriter.WriteElementString("Version", Application.ProductVersion.ToString());
                xmlWriter.WriteStartElement("Options"); // <Options>
                xmlWriter.WriteStartElement("DataSource"); // <DataSource>
                xmlWriter.WriteElementString("Typ", "Spreadsheet");
                xmlWriter.WriteElementString("Parameter", dataSource.Parameter);
                xmlWriter.WriteEndElement(); // </DataSource>
                xmlWriter.WriteStartElement("PDF"); // <PDF>
                xmlWriter.WriteElementString("Filepath", tbPDF.Text);
                xmlWriter.WriteEndElement(); // </PDF>
                xmlWriter.WriteStartElement("Spreadsheet"); // <Spreadsheet>
                xmlWriter.WriteElementString("Table", (string)cbSpreadsheetTable.SelectedItem);
                xmlWriter.WriteEndElement();  // </Spreadsheet>
                
                SincronizaOpciones();
                opt.Save(xmlWriter, dataSource);
                
                xmlWriter.WriteEndElement(); // </Options>

                xmlWriter.WriteStartElement("PDFFieldValues"); // <PDFFieldValues>
                foreach (var pdfField in pdfFields.Values)
                {
                    pdfField.XMLWritePDFField(xmlWriter);
                }
                xmlWriter.WriteEndElement(); // </PDFFieldValues>
                xmlWriter.WriteEndElement(); // </BulkPDF>
                xmlWriter.WriteEndDocument();
                xmlWriter.Close();
            }
        }

        private void bLoadConfiguration_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "BulkPDF Options|*.bulkpdf";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    XDocument xDocument = XDocument.Parse(File.ReadAllText(openFileDialog.FileName));

                    //// Options
                    var xmlOptions = xDocument.Root.Element("Options");
                    // DataSource
                    ResetDataSource();
                    ResetPDF();
                    dataSource = new Spreadsheet();
                    if (!OpenSpreadsheet(Environment.ExpandEnvironmentVariables(xmlOptions.Element("DataSource").Element("Parameter").Value)))
                    {
                        throw new Exception();
                    }
                    cbSpreadsheetTable.SelectedIndex = cbSpreadsheetTable.Items.IndexOf(xmlOptions.Element("Spreadsheet").Element("Table").Value);

                    // PDF
                    if (!OpenPDF(Environment.ExpandEnvironmentVariables(xmlOptions.Element("PDF").Element("Filepath").Value)))
                    {
                        throw new Exception();
                    }

                    //// PDFFieldValues
                    foreach (var node in xDocument.Root.Element("PDFFieldValues").Descendants("PDFFieldValue"))
                    {
                        var name = node.Element("Name").Value;

                        for (int row = 0; row < dgvBulkPDF.Rows.Count; row++)
                        {
                            if ((string)dgvBulkPDF.Rows[row].Cells[0].Value == name)
                            {
                                var pdfField = pdfFields[name];
                                pdfField.Rellena(node);
                                pdfFields[name] = pdfField;

                                if (pdfFields[name].UseValueFromDataSource)
                                {
                                    string value = pdfFields[name].DataSourceValue;
                                    if (pdfFields[name].MakeReadOnly)
                                        value = "[#]" + value;

                                    dgvBulkPDF.Rows[row].Cells["ColValue"].Value = value;
                                }
                                else
                                {
                                    dgvBulkPDF.Rows[row].Cells["ColValue"].Value = pdfFields[name].CurrentValue;
                                }
                            }
                        }
                    }

                    //// Filename

                    //var xmlFilename = xmlOptions.Element("Filename");
                   //opt = Opciones.Load(xmlOptions, dataSource);
                    
                    SincronizaControles(Opciones.Load(xmlOptions, dataSource));
                    
                    wizardPages.SelectedIndex = wizardPages.TabPages.Count - 1;
                }
                catch (Exception ex)
                {
                    ExceptionHandler.Throw(Properties.Resources.ExceptionConfigurationFileIsCorrupted, ex.ToString());
                }
            }
        }
        // Sincroniza los controles con los valores de opciones nuevas
        private void SincronizaControles(Opciones newOpt)
        {
            // Opciones generales
            cRowsPerPag.Value = newOpt.RowsPerPag;
            cGrouped.Checked = newOpt.Grouped;
            cGroupByColumn.SelectedIndex = newOpt.GroupByColumn;                                // cbDataSourceColumnsFilename.Items.IndexOf(xmlFilename.Element("DataSourceGroupBy").Value);
            cbFinalize.Checked = newOpt.Finalize;
            cbUnicode.Checked = newOpt.Unicode;

            cFilter.Checked = opt.Filtered;
            cFilterField.SelectedIndex = newOpt.DSColFilter;
            cFilterCheck.SelectedItem = opt.FilterOperator;

            // Filename
            cPrefix.Text = newOpt.Prefix;                                            // xmlFilename.Element("Prefix").Value;
            cUseValueFromDataSource.Checked = newOpt.UseValueFromDataSource;         // Convert.ToBoolean(xmlFilename.Element("ValueFromDataSource").Value);
            cDataSourceColumnsFilename.SelectedIndex = newOpt.DataSourceColumnsFilename;              // cbDataSourceColumnsFilename.Items.IndexOf(xmlFilename.Element("DataSource").Value);
            cSufix1.Text = newOpt.Sufix1;                                           // xmlFilename.Element("Suffix2").Value;
            cUseGroup.Checked = newOpt.UseGroup;                                    // Convert.ToBoolean(xmlFilename.Element("UseGroup").Value);
            cSufix2.Text = newOpt.Sufix2;                                           // xmlFilename.Element("Suffix2").Value;
            cUseRowNumber.Checked = newOpt.UseRowNumber;                        // Convert.ToBoolean(xmlFilename.Element("RowNumber").Value);
            cUsePag.Checked = newOpt.UsePag;
            tbOutputDir.Text = newOpt.OutputDir;
            cFlatten.Checked = newOpt.Flatten;
            cFusion.Checked = newOpt.Fusion;
            // Custom Font
            cbCustomFont.Checked = newOpt.CustomFont;
            tbCustomFontPath.Text = newOpt.CustomFontPath;
        }
        // Sincroniza las opciones con los controles
        private void SincronizaOpciones()
        {
            // Primero hay que cargar las opcions referentes a agrupación para que se pueda activar el usar grupo
            // Opciones generales
            opt.RowsPerPag = (int)cRowsPerPag.Value;
            opt.Grouped = cGrouped.Checked;
            opt.GroupByColumn = cGroupByColumn.SelectedIndex;                                // cbDataSourceColumnsFilename.Items.IndexOf(xmlFilename.Element("DataSourceGroupBy").Value);
            opt.Finalize = cbFinalize.Checked;
            opt.Unicode = cbUnicode.Checked;

            opt.Filtered = cFilter.Checked;
            opt.DSColFilter = cFilterField.SelectedIndex;
            opt.FilterOperator = (string)cFilterCheck.SelectedItem;

            // Filename
            opt.Fusion = cFusion.Checked;
            opt.Flatten = cFlatten.Checked;
            opt.Prefix = cPrefix.Text;                                            // xmlFilename.Element("Prefix").Value;
            opt.UseValueFromDataSource = cUseValueFromDataSource.Checked;         // Convert.ToBoolean(xmlFilename.Element("ValueFromDataSource").Value);
            opt.DataSourceColumnsFilename = cDataSourceColumnsFilename.SelectedIndex;              // cbDataSourceColumnsFilename.Items.IndexOf(xmlFilename.Element("DataSource").Value);
            opt.Sufix1 = cSufix1.Text;                                           // xmlFilename.Element("Suffix2").Value;
            opt.UseGroup = cUseGroup.Checked;                                  // Convert.ToBoolean(xmlFilename.Element("UseGroup").Value);
            opt.Sufix2 = cSufix2.Text;                                           // xmlFilename.Element("Suffix2").Value;
            opt.UseRowNumber = cUseRowNumber.Checked;                        // Convert.ToBoolean(xmlFilename.Element("RowNumber").Value);
            opt.UsePag = cUsePag.Checked;                                         // Convert.ToBoolean(xmlFilename.Element("UseGroup").Value);
            opt.OutputDir = tbOutputDir.Text;

            // Custom Font
            opt.CustomFont = cbCustomFont.Checked;
            opt.CustomFontPath = tbCustomFontPath.Text;
        }

        private void PopulateFilterOperators()
        {
            cFilterCheck.Items.Clear();

            foreach (string item in PDFFiller.filtros.Keys)
            {
                cFilterCheck.Items.Add(Properties.Resources.ResourceManager.GetString(item, Properties.Resources.Culture));
            }
            cFilterCheck.SelectedIndex = 0;
        }


        /**************************************************/
        #endregion
        /**************************************************/

        /**************************************************/
        #region DataSource
        /**************************************************/

        private void UpdateDataSource()
        {
            // Textbox
            tbSpreadsheet.Text = dataSource.Parameter;
            lPossibleRowsValue.Text = dataSource.PossibleRows.ToString();
            lPossibleColumnsValue.Text = dataSource.Columns.Count.ToString();

            // Dropdown
            cbDataSourceColumnsExampleSpreadsheet.Items.Clear();
            cDataSourceColumnsFilename.Items.Clear();
            cGroupByColumn.Items.Clear();
            cFilterField.Items.Clear();
            foreach (var column in dataSource.Columns)
            {
                cbDataSourceColumnsExampleSpreadsheet.Items.Add(column);
                cDataSourceColumnsFilename.Items.Add(column);
                cGroupByColumn.Items.Add(column);
                cFilterField.Items.Add(column);
            }
            cbDataSourceColumnsExampleSpreadsheet.SelectedIndex = 0;
            cDataSourceColumnsFilename.SelectedIndex = 0;
            cGroupByColumn.SelectedIndex = 0;
            cFilterField.SelectedIndex = 0;
        }

        private void ResetDataSource()
        {
            tbSpreadsheet.Text = "";
            dataSource = null;
            lPossibleRowsValue.Text = "0";
            lPossibleColumnsValue.Text = "0";
            cbDataSourceColumnsExampleSpreadsheet.Items.Clear();
            cDataSourceColumnsFilename.Items.Clear();
            cbSpreadsheetTable.Items.Clear();
        }

        /**************************************************/
        #endregion
        /**************************************************/

        /**************************************************/
        #region ExampleFilename
        /**************************************************/

        /// <summary>
        /// Evento para cambios en checks que afectan a la presentacion
        /// como cFusion y  cGrouped
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cCheckOptionsChanged(object sender, EventArgs e)
        {
            cGroupByColumn.Enabled = cGrouped.Checked;
            cUseGroup.Checked = cUseGroup.Checked && cGrouped.Checked;

            bool nf = !cFusion.Checked;
            if (cFusion.Checked)
            {
                cFlatten.Checked = true;
            }
            cFlatten.Enabled = nf;
            cUseValueFromDataSource.Enabled = nf;
            cDataSourceColumnsFilename.Enabled = nf;
            cSufix1.Enabled = nf;
            cUseGroup.Enabled = nf && cGrouped.Checked;
            cSufix2.Enabled = nf;
            cUseRowNumber.Enabled = nf;
            cUsePag.Enabled = nf;
            UpdateExampleFilename();
        }

        private void cFilenameChanged(object sender, EventArgs e)
        {
            UpdateExampleFilename();
        }

        private void UpdateExampleFilename()
        {
            SincronizaOpciones();
            //dataSource.ResetRowCounter();
            tbExampleFilename.Text = cFusion.Checked ? cPrefix.Text+".pdf": ConcatFilename(1);
        }

        /// <summary>
        /// Cambia la disponibilidad de algunos controles cuando cambia el check de cFilter
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cFilter_CheckedChanged(object sender, EventArgs e)
        {
            cFilterField.Enabled = cFilterCheck.Enabled = cFilter.Checked;
        }

        /**************************************************/
        #endregion
        /**************************************************/

        /**************************************************/
        #region SpreadsheetSelect
        /**************************************************/

        private void bSelectSpreadsheet_Click(object sender, EventArgs e)
        {
            // Select File
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "Spreadsheet|*.xlsx;*.xlsm";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // DataSource
                if (dataSource != null)
                    dataSource.Close();
                ResetDataSource();
                ResetPDF();
                dataSource = new Spreadsheet();

                OpenSpreadsheet(openFileDialog.FileName);
            }
        }

        private bool OpenSpreadsheet(string filePath)
        {
            try
            {
                dataSource.Open(filePath);

                // Sheet
                var sheetNames = ((Spreadsheet)dataSource).GetSheetNames();
                foreach (var sheet in sheetNames)
                    cbSpreadsheetTable.Items.Add(sheet);
                cbSpreadsheetTable.SelectedIndex = 0;

                UpdateDataSource();
                return true;
            }
            catch (IOException ex)
            {
                ExceptionHandler.Throw(Properties.Resources.ExceptionSpreadsheetAlreadyInUse, ex.ToString());
            }
            catch (FileFormatException ex)
            {
                ExceptionHandler.Throw(Properties.Resources.ExceptionSpreadsheetIsCorrupted, ex.ToString());
            }
            return false;
        }

        private void cbTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ((Spreadsheet)dataSource).SetSheet((string)cbSpreadsheetTable.SelectedItem);
            UpdateDataSource();
            ResetPDF();
        }

        /**************************************************/
        #endregion
        /**************************************************/


        private void bShortcutCreator_Click(object sender, EventArgs e)
        {
            (new ShortcutCreator()).ShowDialog();
        }

        private void llSupport_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("mailto:support@bulkpdf.de");
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            var donateForm = new DonateForm();
            donateForm.ShowDialog();
        }

        private void bSelectOwnFont_Click(object sender, EventArgs e)
        {
            // Select File
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "Font|*.ttf;";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                tbCustomFontPath.Text = openFileDialog.FileName;
            }
        }
    }
}
