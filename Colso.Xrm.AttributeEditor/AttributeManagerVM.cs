using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.ServiceModel;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using Colso.Xrm.AttributeEditor.Annotations;
using Colso.Xrm.AttributeEditor.AppCode;
using Colso.Xrm.AttributeEditor.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using NuGet;

namespace Colso.Xrm.AttributeEditor
{
    public class AttributeManagerVM : INotifyPropertyChanged
    {
        private readonly IExcelDocumentProcessor _excelDocumentProcessor;
        public Func<IOrganizationService> GetService { get; }
        private bool _busy;
        private string _busyMessage;
        private string _selectedEntity;
        private string _templateForUpload;
        private string _statusMessage;
        private bool _create;
        private bool _update;
        private bool _delete;

        public ObservableCollection<string> Entities { get; }
        public ObservableCollection<ListViewItem> Attributes { get; }
        public ObservableCollection<string> ColumnHeadings { get; }

        public bool Busy
        {
            get { return _busy; }
            private set
            {
                if (value == _busy) return;
                _busy = value;
                OnPropertyChanged();
            }
        }

        public string BusyMessage
        {
            get { return _busyMessage; }
            private set
            {
                if (value == _busyMessage) return;
                _busyMessage = value;
                OnPropertyChanged();
            }
        }
        public string SelectedEntity
        {
            get { return _selectedEntity; }
            set
            {
                if (value == _selectedEntity) return;
                _selectedEntity = value;
                PopulateAttributes();
                OnPropertyChanged();
            }
        }
        public string TemplateForUpload
        {
            get { return _templateForUpload; }
            set
            {
                if (value == _templateForUpload) return;
                _templateForUpload = value;
                OnPropertyChanged();
            }
        }
        public string StatusMessage
        {
            get { return _statusMessage; }
            private set
            {
                if (value == _statusMessage) return;
                _statusMessage = value;
                OnPropertyChanged();
            }
        }
        public bool Create
        {
            get { return _create; }
            set
            {
                if (value == _create) return;
                _create = value;
                OnPropertyChanged();
            }
        }
        public bool Update
        {
            get { return _update; }
            set
            {
                if (value == _update) return;
                _update = value;
                OnPropertyChanged();
            }
        }
        public bool Delete
        {
            get { return _delete; }
            set
            {
                if (value == _delete) return;
                _delete = value;
                OnPropertyChanged();
            }
        }

        public ICommand LoadEntitiesCommand => new RelayCommand(p => LoadEntities(), CanExecute);
        public ICommand DownloadTemplateCommand => new RelayCommand(p => DownloadTemplate(), CanExecute);
        public ICommand SelectTemplateForUploadCommand => new RelayCommand(p => SelectTemplateForUpload(), CanExecute);
        public ICommand UploadTemplateCommand => new RelayCommand(p => UploadTemplate(), CanExecute);
        public ICommand SaveChangesCommand => new RelayCommand(p => SaveChanges(), CanExecute);

        public AttributeManagerVM(Func<IOrganizationService> serviceProvider, IExcelDocumentProcessor excelDocumentProcessor)
        {
            _excelDocumentProcessor = excelDocumentProcessor;
            GetService = serviceProvider;

            Entities = new ObservableCollection<string>();
            Attributes = new ObservableCollection<ListViewItem>();
            ColumnHeadings = new ObservableCollection<string>();

            Create = true;
            Update = true;
            Delete = false;
        }

        private bool CanExecute(object parameter)
        {
            return !_busy;
        }

        private void LoadEntities()
        {
            Entities.Clear();

            Busy = true;
            BusyMessage = "Retrieving Entities";

            Task.Factory.StartNew(() =>
                {
                    return MetadataHelper.RetrieveEntities(GetService()).Select(x => x.LogicalName);
                })
                .ContinueWith(result =>
                {
                    var sortedEntities = result.Result.OrderBy(x => x);

                    Busy = false;

                    Entities.AddRange(sortedEntities);

                    Busy = false;
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void PopulateAttributes()
        {
            Attributes.Clear();

            if (string.IsNullOrEmpty(SelectedEntity))
                return;

            ColumnHeadings.Clear();
            ColumnHeadings.Add("Action");

            foreach (var column in AttributeMetadataRow.GetColumns())
                ColumnHeadings.Add(column.Header);

            Busy = true;
            BusyMessage = "Loading Entity";

            Task.Factory.StartNew(() =>
                {
                    // Retrieve 
                    var entity = MetadataHelper.RetrieveEntity(SelectedEntity, GetService());

                    // Prepare list of items
                    var itemList = new List<ListViewItem>();
                    var orderedAttributes = entity.Attributes.OrderBy(x => x.LogicalName);

                    foreach (AttributeMetadata att in orderedAttributes)
                    {
                        if (att.IsValidForUpdate.Value && att.IsCustomizable.Value)
                        {
                            var attribute = EntityHelper.GetAttributeFromTypeName(att.AttributeType.Value.ToString());

                            if (attribute == null)
                                continue;

                            attribute.LoadFromAttributeMetadata(att);

                            var row = attribute.ToAttributeMetadataRow();

                            var item = row.ToListViewItem();

                            item.Tag = att;

                            itemList.Add(item);
                        }
                    }

                    return itemList;
                })
                .ContinueWith(result =>
                {
                    Attributes.AddRange(result.Result);
                    Busy = false;
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void DownloadTemplate()
        {
            if (SelectedEntity == null)
            {
                MessageBox.Show("You must select an entity", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Title = "Save the entity template",
                Filter = "Excel Workbook|*.xlsx",
                FileName = "attributeeditor.xlsx"
            };
            saveFileDialog.ShowDialog();

            if (string.IsNullOrEmpty(saveFileDialog.FileName))
                return;

            Busy = true;
            BusyMessage = "Exporting Attributes";

            var attributes = Attributes.Select(x => x.Tag as AttributeMetadata);

            Task.Factory.StartNew(() =>
                {
                    var rows = new List<ExcelDocumentProcessor.WorkbookRow>();
                    
                    var headerCells = new List<ExcelDocumentProcessor.WorkbookCell>();
                    foreach (var field in AttributeMetadataRow.GetColumns())
                        headerCells.Add(new ExcelDocumentProcessor.WorkbookCell
                        {
                            Type = (ExcelDocumentProcessor.CellTypes) field.Type,
                            Value = field.Header
                        });
                    rows.Add(new ExcelDocumentProcessor.WorkbookRow { Cells = headerCells });
                    
                    foreach (var attribute in attributes)
                    {
                        if (attribute == null)
                            continue;

                        var attributeType =
                            EntityHelper.GetAttributeFromTypeName(attribute.AttributeType.Value.ToString());

                        if (attributeType == null)
                            continue;

                        attributeType.LoadFromAttributeMetadata(attribute);

                        var metadataRow = attributeType.ToAttributeMetadataRow();
                        
                        rows.Add(metadataRow.ToTableRow());
                    }

                    var sheet = new ExcelDocumentProcessor.WorkbookSheet
                    {
                        Rows = rows
                    };
                    
                    _excelDocumentProcessor.SaveWorkbook(saveFileDialog.FileName, SelectedEntity, sheet);
                })
                .ContinueWith((result) =>
                {
                    Busy = false;
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void SelectTemplateForUpload()
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                TemplateForUpload = file.FileName;
            }
        }

        private void UploadTemplate()
        {
            if (string.IsNullOrEmpty(SelectedEntity))
            {
                MessageBox.Show("You must select an entity", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(TemplateForUpload))
            {
                MessageBox.Show("You must select a template file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var callerSynchronicationContext = TaskScheduler.FromCurrentSynchronizationContext();

            Task.Factory.StartNew(() =>
                {

                    var attributes = Attributes;
                    var errors = new List<Tuple<string, string>>();
                    var nrOfCreates = 0;
                    var nrOfUpdates = 0;
                    var nrOfDeletes = 0;

                    var items = new List<ListViewItem>();

                    var sheet = _excelDocumentProcessor.Read(TemplateForUpload, SelectedEntity);
                    var attrDict = sheet.Rows.ToDictionary(x => x.Cells.First().Value.ToLower());


                    try
                    {
                        // Check all existing attributes
                        foreach (var item in attributes)
                        {
                            var itemLogicalName = item.SubItems[1].Text.ToLower();

                            var row = attrDict.ContainsKey(itemLogicalName)
                                ? attrDict[itemLogicalName]
                                : null;

                            var attributeMetadataRow = AttributeMetadataRow.FromListViewItem(item);

                            attributeMetadataRow.UpdateFromTemplateRow(row);
                            var listViewItem = attributeMetadataRow.ToListViewItem();
                            listViewItem.Tag = item.Tag;

                            items.Add(listViewItem);

                            if (attributeMetadataRow.Action == "Update")
                                nrOfUpdates++;
                            else if (attributeMetadataRow.Action == "Delete")
                                nrOfDeletes++;

                            BusyMessage = $"Processing new attribute {item.Text}...";
                        }

                        // Check new attributes
                        for (int i = 1; i < sheet.Rows.Count(); i++)
                        {
                            var row = sheet.Rows.Skip(i).First();
                            var logicalname = row.Cells.First().Value.ToLower();

                            if (!string.IsNullOrEmpty(logicalname))
                            {
                                var attribute = attributes
                                    .Select(a => (AttributeMetadata)a.Tag).FirstOrDefault(a => a.LogicalName == logicalname);

                                if (attribute == null)
                                {
                                    BusyMessage = string.Format("Processing new attribute {0}...", logicalname);

                                    var attributeMetadataRow =
                                        AttributeMetadataRow.FromTableRow(row);
                                    attributeMetadataRow.Action = "Create";

                                    items.Add(attributeMetadataRow.ToListViewItem());

                                    nrOfCreates++;
                                }
                            }
                        }

                        StatusMessage = $"{nrOfCreates} create; {nrOfUpdates} update; {nrOfDeletes} delete";;
                    }
                    catch (FaultException<OrganizationServiceFault> error)
                    {
                        errors.Add(new Tuple<string, string>(SelectedEntity, error.Message));
                        StatusMessage = string.Empty;
                    }

                    return new {Items = items, Errors = errors};
                })
                .ContinueWith((result) =>
                {
                    Attributes.Clear();
                    Attributes.AddRange(result.Result.Items);

                    Busy = false;

                    var errors = result.Result.Errors;

                    if (errors.Count > 0)
                    {
                        var errorDialog = new ErrorList(result.Result.Errors);
                        errorDialog.ShowDialog();
                    }
                }, callerSynchronicationContext);
        }

        private void SaveChanges()
        {
            if (string.IsNullOrEmpty(SelectedEntity))
            {
                MessageBox.Show("You must select an entity", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Busy = true;
            BusyMessage = "Processing entity...";

            var itemsToProcess = Attributes;

            Task.Factory.StartNew(() =>
                {
                    var errors = new List<Tuple<string, string>>();
                    var helper = new EntityHelper(SelectedEntity, 1033, GetService());

                    foreach (var item in itemsToProcess)
                    {
                        // TODO: Validate each row
                        //if (item.SubItems.Count == 6)
                        //{ 
                        var action = item.SubItems[0].Text;

                        if (!string.IsNullOrEmpty(action) && !string.IsNullOrEmpty(item.Text))
                        {
                            var attributeMetadata = AttributeMetadataRow.FromListViewItem(item);

                            BusyMessage = $"Processing attribute {attributeMetadata.LogicalName}...";

                            try
                            {
                                var attribute = EntityHelper.GetAttributeFromTypeName(attributeMetadata.AttributeType);

                                if (attribute == null)
                                {
                                    errors.Add(new Tuple<string, string>(item.Text,
                                        $"The Attribute Type \"{attributeMetadata.AttributeType}\" is not yet supported."));
                                    continue;
                                }

                                attribute.LoadFromAttributeMetadataRow(attributeMetadata);
                                attribute.Entity = SelectedEntity;

                                switch (action)
                                {
                                    case "Create":
                                        if (Create) attribute.CreateAttribute(GetService());
                                        break;
                                    case "Update":
                                        if (Update) attribute.UpdateAttribute(GetService());
                                        break;
                                    case "Delete":
                                        if (Delete) attribute.DeleteAttribute(GetService());
                                        break;
                                }
                            }
                            catch (FaultException<OrganizationServiceFault> error)
                            {
                                errors.Add(new Tuple<string, string>(item.Text, error.Message));
                            }
                        }
                    }

                    helper.Publish();

                    return errors;
                })
                .ContinueWith((result) =>
                {
                    Busy = false;
                    StatusMessage = string.Empty;

                    var errors = result.Result;

                    if (errors.Count > 0)
                    {
                        var errorDialog = new ErrorList(result.Result);
                        errorDialog.ShowDialog();
                    }

                    PopulateAttributes();
                });
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
