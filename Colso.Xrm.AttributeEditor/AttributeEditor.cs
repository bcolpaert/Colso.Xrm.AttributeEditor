using Colso.Xrm.AttributeEditor.AppCode;
using Colso.Xrm.AttributeEditor.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.ServiceModel;
using System.Windows.Forms;
using System.Xml;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Args;
using XrmToolBox.Extensibility.Interfaces;

namespace Colso.Xrm.AttributeEditor
{
    public partial class AttributeEditor : UserControl, IXrmToolBoxPluginControl, IGitHubPlugin, IHelpPlugin, IStatusBarMessenger
    {
        #region Variables

        /// <summary>
        /// Information panel
        /// </summary>
        private Panel informationPanel;

        /// <summary>
        /// Dynamics CRM 2011 organization service
        /// </summary>
        private IOrganizationService service;

        private bool workingstate = false;
        private Dictionary<string, int> lvSortcolumns = new Dictionary<string, int>();

        private System.Drawing.Color ColorCreate = System.Drawing.Color.Green;
        private System.Drawing.Color ColorUpdate = System.Drawing.Color.Blue;
        private System.Drawing.Color ColorDelete = System.Drawing.Color.Red;

        #endregion Variables

        public AttributeEditor()
        {
            InitializeComponent();
        }

        #region XrmToolbox

        public event EventHandler OnCloseTool;
        public event EventHandler OnRequestConnection;
        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;

        public Image PluginLogo
        {
            get { return null; }
        }

        public IOrganizationService Service
        {
            get { throw new NotImplementedException(); }
        }

        public string HelpUrl
        {
            get
            {
                return "https://github.com/MscrmTools/Colso.Xrm.DataTransporter/wiki";
            }
        }

        public string RepositoryName
        {
            get
            {
                return "Colso.Xrm.AttributeEditor";
            }
        }

        public string UserName
        {
            get
            {
                return "MscrmTools";
            }
        }

        public void ClosingPlugin(PluginCloseInfo info)
        {
            if (info.FormReason != CloseReason.None ||
                info.ToolBoxReason == ToolBoxCloseReason.CloseAll ||
                info.ToolBoxReason == ToolBoxCloseReason.CloseAllExceptActive)
            {
                return;
            }

            info.Cancel = MessageBox.Show(@"Are you sure you want to close this tab?", @"Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes;
        }

        public void UpdateConnection(IOrganizationService newService, ConnectionDetail connectionDetail, string actionName = "", object parameter = null)
        {
            service = newService;
        }

        public string GetCompany()
        {
            return GetType().GetCompany();
        }

        public string GetMyType()
        {
            return GetType().FullName;
        }

        public string GetVersion()
        {
            return GetType().Assembly.GetName().Version.ToString();
        }

        #endregion XrmToolbox

        #region Form events

        private void tsbCloseThisTab_Click(object sender, EventArgs e)
        {
            if (OnCloseTool != null)
                OnCloseTool(this, null);
        }

        private void tsbLoadEntities_Click(object sender, EventArgs e)
        {
            if (service == null)
            {
                if (OnRequestConnection != null)
                {
                    var args = new RequestConnectionEventArgs
                    {
                        ActionName = "Load",
                        Control = this
                    };
                    OnRequestConnection(this, args);
                }
            }
            else
            {
                PopulateEntities();
            }
        }

        private void tsbPublish_Click(object sender, EventArgs e)
        {
            SaveChanges();
        }

        private void cmbEntities_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateAttributes();
        }

        private void lvAttributes_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            SetListViewSorting(lvAttributes, e.Column);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            ImportAttributes();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ExportAttributes();
        }

        private void btnSelectTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                txtTemplatePath.Text = file.FileName;
            }
        }

        #endregion Form events

        #region Methods

        private void ManageWorkingState(bool working)
        {
            workingstate = working;
            cmbEntities.Enabled = !working;
            gbSettings.Enabled = !working;
            gbAttributes.Enabled = !working;
            Cursor = working ? Cursors.WaitCursor : Cursors.Default;
        }

        private void PopulateEntities()
        {
            if (!workingstate)
            {
                // Reinit other controls
                cmbEntities.Items.Clear();
                ManageWorkingState(true);

                informationPanel = InformationPanel.GetInformationPanel(this, "Loading entities...", 340, 150);

                // Launch treatment
                var bwFill = new BackgroundWorker();
                bwFill.DoWork += (sender, e) =>
                {
                    // Retrieve 
                    List<EntityMetadata> sourceList = MetadataHelper.RetrieveEntities(service);

                    // Prepare list of items
                    var sourceEntitiesList = new List<EntityItem>();

                    foreach (EntityMetadata entity in sourceList)
                        sourceEntitiesList.Add(new EntityItem(entity));

                    e.Result = sourceEntitiesList.OrderBy(i => i.DisplayName).ToArray();
                };
                bwFill.RunWorkerCompleted += (sender, e) =>
                {
                    informationPanel.Dispose();

                    if (e.Error != null)
                    {
                        MessageBox.Show(this, "An error occured: " + e.Error.Message, "Error", MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                    }
                    else
                    {
                        var items = (EntityItem[])e.Result;
                        if (items.Length == 0)
                        {
                            MessageBox.Show(this, "The system does not contain any entities", "Warning", MessageBoxButtons.OK,
                                            MessageBoxIcon.Warning);
                        }
                        else
                        {
                            cmbEntities.DisplayMember = "DisplayName";
                            cmbEntities.ValueMember = "SchemaName";
                            cmbEntities.Items.AddRange(items);
                        }
                    }

                    ManageWorkingState(false);
                };
                bwFill.RunWorkerAsync();
            }
        }

        private void PopulateAttributes()
        {
            if (!workingstate)
            {
                // Reinit other controls
                lvAttributes.Items.Clear();

                if (cmbEntities.SelectedItem != null)
                {
                    var entityitem = (EntityItem)cmbEntities.SelectedItem;

                    if (!string.IsNullOrEmpty(entityitem.LogicalName))
                    {
                        ManageWorkingState(true);

                        // Launch treatment
                        var bwFill = new BackgroundWorker();
                        bwFill.DoWork += (sender, e) =>
                        {
                            // Retrieve 
                            var entity = MetadataHelper.RetrieveEntity(entityitem.LogicalName, service);

                            // Prepare list of items
                            var itemList = new List<ListViewItem>();

                            foreach (AttributeMetadata att in entity.Attributes)
                            {
                                if (att.IsValidForUpdate.Value && att.IsCustomizable.Value)
                                {
                                    var name = att.DisplayName == null ?
                                        string.Empty :
                                        att.DisplayName.UserLocalizedLabel == null ?
                                        att.DisplayName.LocalizedLabels.Select(l => l.Label).FirstOrDefault() :
                                        att.DisplayName.UserLocalizedLabel.Label;
                                    var item = new ListViewItem(att.LogicalName);
                                    item.Tag = att;
                                    item.SubItems.Add(name);
                                    item.SubItems.Add(att.AttributeType.Value.ToString());
                                    item.SubItems.Add(att.IsCustomAttribute.Value ? "Unmanaged" : "Managed");
                                    item.SubItems.Add(att.RequiredLevel.Value.ToString());
                                    item.SubItems.Add(string.Empty);
                                    itemList.Add(item);
                                }
                            }

                            e.Result = itemList;
                        };
                        bwFill.RunWorkerCompleted += (sender, e) =>
                        {
                            if (e.Error != null)
                            {
                                MessageBox.Show(this, "An error occured: " + e.Error.Message, "Error", MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                            }
                            else
                            {
                                var items = (List<ListViewItem>)e.Result;
                                if (items.Count == 0)
                                {
                                    MessageBox.Show(this, "The entity does not contain any attributes", "Warning", MessageBoxButtons.OK,
                                                    MessageBoxIcon.Warning);
                                }
                                else
                                {
                                    lvAttributes.Items.AddRange(items.ToArray());
                                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(string.Format("{0} attributes loaded", items.Count)));
                                }
                            }

                            ManageWorkingState(false);
                        };
                        bwFill.RunWorkerAsync();
                    }
                }
            }
        }

        private void ExportAttributes()
        {
            if (cmbEntities.SelectedItem == null)
            {
                MessageBox.Show("You must select an entity", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ManageWorkingState(true);

            var saveFileDlg = new SaveFileDialog()
            {
                Title = "Save the entity template",
                Filter = "Excel Workbook|*.xlsx",
                FileName = "attributeeditor.xlsx"
            };
            saveFileDlg.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (!string.IsNullOrEmpty(saveFileDlg.FileName))
            {
                informationPanel = InformationPanel.GetInformationPanel(this, "Exporting attributes...", 340, 150);
                SendMessageToStatusBar(this, new StatusBarMessageEventArgs("Initializing attributes..."));

                var bwTransferData = new BackgroundWorker { WorkerReportsProgress = true };
                bwTransferData.DoWork += (sender, e) =>
                {
                    var worker = (BackgroundWorker)sender;
                    var attributes = (List<AttributeMetadata>)e.Argument;
                    var entityitem = (EntityItem)cmbEntities.SelectedItem;
                    var errors = new List<Tuple<string, string>>();

                    try
                    {
                        using (SpreadsheetDocument document = SpreadsheetDocument.Create(saveFileDlg.FileName, SpreadsheetDocumentType.Workbook))
                        {
                            WorkbookPart workbookPart;
                            Sheets sheets;
                            SheetData sheetData;
                            TemplateHelper.InitDocument(document, out workbookPart, out sheets);
                            TemplateHelper.CreateSheet(workbookPart, sheets, entityitem.LogicalName, out sheetData);

                            #region Header

                            var headerRow = new Row();

                            headerRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, "Logical Name"));
                            headerRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, "Display Name"));
                            headerRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, "Type"));
                            headerRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, "Field Requirement"));

                            sheetData.AppendChild(headerRow);

                            #endregion

                            #region Data

                            foreach (var attribute in attributes)
                            {
                                if (attribute != null)
                                {
                                    var newRow = new Row();

                                    newRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, attribute.LogicalName));
                                    newRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, attribute.DisplayName.UserLocalizedLabel.Label));
                                    newRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, attribute.AttributeType.Value.ToString()));
                                    newRow.AppendChild(TemplateHelper.CreateCell(CellValues.String, attribute.RequiredLevel.Value.ToString()));

                                    sheetData.AppendChild(newRow);
                                }
                            }

                            #endregion

                            workbookPart.Workbook.Save();
                        }
                    }
                    catch (FaultException<OrganizationServiceFault> error)
                    {
                        errors.Add(new Tuple<string, string>(entityitem.LogicalName, error.Message));
                    }

                    e.Result = errors;
                };
                bwTransferData.RunWorkerCompleted += (sender, e) =>
                {
                    Controls.Remove(informationPanel);
                    informationPanel.Dispose();
                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(string.Empty));
                    ManageWorkingState(false);

                    var errors = (List<Tuple<string, string>>)e.Result;

                    if (errors.Count > 0)
                    {
                        var errorDialog = new ErrorList((List<Tuple<string, string>>)e.Result);
                        errorDialog.ShowDialog(ParentForm);
                    }
                };
                bwTransferData.ProgressChanged += (sender, e) =>
                {
                    InformationPanel.ChangeInformationPanelMessage(informationPanel, e.UserState.ToString());
                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(e.UserState.ToString()));
                };
                bwTransferData.RunWorkerAsync(lvAttributes.Items.Cast<ListViewItem>().Select(v => (AttributeMetadata)v.Tag).ToList());
            }
        }

        private void ImportAttributes()
        {
            if (cmbEntities.SelectedItem == null)
            {
                MessageBox.Show("You must select an entity", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtTemplatePath.Text))
            {
                MessageBox.Show("You must select a template file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ManageWorkingState(true);

            informationPanel = InformationPanel.GetInformationPanel(this, "Loading template...", 340, 150);
            SendMessageToStatusBar(this, new StatusBarMessageEventArgs("Loading template..."));

            var bwTransferData = new BackgroundWorker { WorkerReportsProgress = true };
            bwTransferData.DoWork += (sender, e) =>
            {
                var worker = (BackgroundWorker)sender;
                var attributes = (List<ListViewItem>)e.Argument;
                var entityitem = (EntityItem)cmbEntities.SelectedItem;
                var errors = new List<Tuple<string, string>>();
                var nrOfCreates = 0;
                var nrOfUpdates = 0;
                var nrOfDeletes = 0;

                try
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(txtTemplatePath.Text, false))
                    {
                        var sheet = document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Where(s => s.Name == entityitem.LogicalName).FirstOrDefault();
                        if (sheet == null)
                        {
                            errors.Add(new Tuple<string, string>(entityitem.LogicalName, "Entity sheet is not found in the template!"));
                        }
                        else
                        {
                            var sheetid = sheet.Id;
                            var sheetData = ((WorksheetPart)document.WorkbookPart.GetPartById(sheetid)).Worksheet.GetFirstChild<SheetData>();
                            var templaterows = sheetData.ChildElements.Cast<Row>().ToArray();
                            // For shared strings, look up the value in the shared strings table.
                            var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;

                            lvAttributes.BeginUpdate();
                            // Check all existing attributes
                            foreach (var item in attributes)
                            {
                                var row = templaterows.Where(r => TemplateHelper.GetCellValue(r, 0, stringTable) == item.Text).FirstOrDefault();

                                InformationPanel.ChangeInformationPanelMessage(informationPanel, string.Format("Processing new attribute {0}...", item.Text));
                                if (row != null)
                                {
                                    if (IsChanged(item, row, stringTable))
                                    {
                                        nrOfUpdates++;
                                        item.UseItemStyleForSubItems = false;
                                        item.SubItems[5].Text = "Update";
                                        item.SubItems[5].ForeColor = ColorUpdate;
                                    }
                                }
                                else if (item.SubItems[3].Text != "Managed")
                                {
                                    nrOfDeletes++;
                                    item.UseItemStyleForSubItems = true;
                                    item.SubItems[5].Text = "Delete";
                                    item.ForeColor = ColorDelete;
                                }
                            }

                            // Check new attributes
                            for (int i = 1; i < templaterows.Length; i++)
                            {
                                var row = templaterows[i];
                                var logicalname = TemplateHelper.GetCellValue(row, 0, stringTable);

                                if (!string.IsNullOrEmpty(logicalname))
                                {
                                    var attribute = attributes.Select(a => (AttributeMetadata)a.Tag).Where(a => a.LogicalName == logicalname).FirstOrDefault();

                                    if (attribute == null)
                                    {
                                        InformationPanel.ChangeInformationPanelMessage(informationPanel, string.Format("Processing new attribute {0}...", logicalname));

                                        var displayname = TemplateHelper.GetCellValue(row, 1, stringTable);
                                        var type = TemplateHelper.GetCellValue(row, 2, stringTable);
                                        var requirement = TemplateHelper.GetCellValue(row, 3, stringTable);

                                        var item = new ListViewItem(logicalname);
                                        item.ForeColor = ColorCreate;
                                        item.SubItems.Add(string.IsNullOrEmpty(displayname) ? logicalname : displayname);
                                        item.SubItems.Add(string.IsNullOrEmpty(type) ? "String" : type);
                                        item.SubItems.Add("Unmanaged");
                                        item.SubItems.Add(string.IsNullOrEmpty(requirement) ? "None" : requirement);
                                        item.SubItems.Add("Create");
                                        lvAttributes.Items.Add(item);

                                        nrOfCreates++;
                                    }
                                }
                            }

                            lvAttributes.EndUpdate();
                        }
                    }
                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(string.Format("{0} create; {1} update; {2} delete", nrOfCreates, nrOfUpdates, nrOfDeletes)));
                }
                catch (FaultException<OrganizationServiceFault> error)
                {
                    errors.Add(new Tuple<string, string>(entityitem.LogicalName, error.Message));
                    lvAttributes.EndUpdate();
                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(string.Empty));
                }

                e.Result = errors;
            };
            bwTransferData.RunWorkerCompleted += (sender, e) =>
            {
                Controls.Remove(informationPanel);
                informationPanel.Dispose();
                ManageWorkingState(false);

                var errors = (List<Tuple<string, string>>)e.Result;

                if (errors.Count > 0)
                {
                    var errorDialog = new ErrorList((List<Tuple<string, string>>)e.Result);
                    errorDialog.ShowDialog(ParentForm);
                }
            };
            bwTransferData.ProgressChanged += (sender, e) =>
            {
                InformationPanel.ChangeInformationPanelMessage(informationPanel, e.UserState.ToString());
                SendMessageToStatusBar(this, new StatusBarMessageEventArgs(e.UserState.ToString()));
            };
            bwTransferData.RunWorkerAsync(lvAttributes.Items.Cast<ListViewItem>().ToList());
        }

        private void SaveChanges()
        {
            if (cmbEntities.SelectedItem == null)
            {
                MessageBox.Show("You must select an entity", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ManageWorkingState(true);
            informationPanel = InformationPanel.GetInformationPanel(this, "Processing entity...", 340, 150);

            var bwTransferData = new BackgroundWorker { WorkerReportsProgress = true };
            bwTransferData.DoWork += (sender, e) =>
            {
                var worker = (BackgroundWorker)sender;
                var entityitem = (EntityItem)cmbEntities.SelectedItem;
                var errors = new List<Tuple<string, string>>();
                var helper = new EntityHelper(entityitem.LogicalName, entityitem.LanguageCode, service);

                foreach (var item in lvAttributes.Items.Cast<ListViewItem>())
                {
                    if (item.SubItems.Count == 6)
                    { 
                        var action = item.SubItems[5].Text;

                        if (!string.IsNullOrEmpty(action) && !string.IsNullOrEmpty(item.Text))
                        {
                            InformationPanel.ChangeInformationPanelMessage(informationPanel, string.Format("Processing attribute {0}...", item.Text));

                            try
                            {
                                switch (action)
                                {
                                    case "Create":
                                        if (cbCreate.Checked) helper.CreateAttribute(item.Text, item.SubItems[1].Text, item.SubItems[2].Text, item.SubItems[4].Text);
                                        break;
                                    case "Update":
                                        if (cbUpdate.Checked) helper.UpdateAttribute(item.Text, item.SubItems[1].Text, item.SubItems[4].Text);
                                        break;
                                    case "Delete":
                                        if (cbDelete.Checked) helper.DeleteAttribute(item.Text);
                                        break;
                                }
                            }
                            catch (FaultException<OrganizationServiceFault> error)
                            {
                                errors.Add(new Tuple<string, string>(item.Text, error.Message));
                            }
                        }
                    } else
                    {
                        errors.Add(new Tuple<string, string>(item.Name, string.Format("Invalid row! Unexpected subitems count [{0}]", item.SubItems.Count)));
                    }
                }

                helper.Publish();
                e.Result = errors;
            };
            bwTransferData.RunWorkerCompleted += (sender, e) =>
            {
                Controls.Remove(informationPanel);
                informationPanel.Dispose();
                SendMessageToStatusBar(this, new StatusBarMessageEventArgs(string.Empty));
                ManageWorkingState(false);

                var errors = (List<Tuple<string, string>>)e.Result;

                if (errors.Count > 0)
                {
                    var errorDialog = new ErrorList((List<Tuple<string, string>>)e.Result);
                    errorDialog.ShowDialog(ParentForm);
                }

                PopulateAttributes();
            };
            bwTransferData.ProgressChanged += (sender, e) =>
            {
                InformationPanel.ChangeInformationPanelMessage(informationPanel, e.UserState.ToString());
                SendMessageToStatusBar(this, new StatusBarMessageEventArgs(e.UserState.ToString()));
            };
            bwTransferData.RunWorkerAsync();
        }

        private void Entity_OnStatusMessage(object sender, EventArgs e)
        {
            SendMessageToStatusBar(this, new StatusBarMessageEventArgs(((StatusMessageEventArgs)e).Message));
        }

        private void SetListViewSorting(ListView listview, int column)
        {
            int currentSortcolumn = -1;
            if (lvSortcolumns.ContainsKey(listview.Name))
                currentSortcolumn = lvSortcolumns[listview.Name];
            else
                lvSortcolumns.Add(listview.Name, currentSortcolumn);

            if (currentSortcolumn != column)
            {
                lvSortcolumns[listview.Name] = column;
                listview.Sorting = SortOrder.Ascending;
            }
            else
            {
                if (listview.Sorting == SortOrder.Ascending)
                    listview.Sorting = SortOrder.Descending;
                else
                    listview.Sorting = SortOrder.Ascending;
            }

            listview.ListViewItemSorter = new ListViewItemComparer(column, listview.Sorting);
        }

        private bool IsChanged(ListViewItem item, Row row, SharedStringTable stringTable)
        {
            var changed = false;

            var displayName = TemplateHelper.GetCellValue(row, 1, stringTable);
            if (item.SubItems[1].Text != displayName && !string.IsNullOrEmpty(displayName))
            {
                item.SubItems[1].Text = displayName;
                item.SubItems[1].ForeColor = ColorUpdate;
                changed = true;
            }

            var requiredLevel = TemplateHelper.GetCellValue(row, 4, stringTable);
            if (item.SubItems[4].Text != requiredLevel && !string.IsNullOrEmpty(requiredLevel))
            {
                item.SubItems[4].Text = requiredLevel;
                item.SubItems[4].ForeColor = ColorUpdate;
                changed = true;
            }

            return changed;
        }

        #endregion Methods

    }
}