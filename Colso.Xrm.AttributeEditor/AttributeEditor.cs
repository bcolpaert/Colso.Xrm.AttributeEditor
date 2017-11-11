using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        private IOrganizationService _service;
        private readonly AttributeManagerVM _viewModel;

        #endregion Variables

        public AttributeEditor()
        {
            InitializeComponent();

            Panel informationPanel = null;

            _viewModel = new AttributeManagerVM(() => _service, new ExcelDocumentProcessor());

            var uiSynchronizationContext = TaskScheduler.FromCurrentSynchronizationContext();

            // Bind form elements to View Model
            _viewModel.PropertyChanged += (sender, args) =>
            {
                Task.Factory.StartNew(() =>
                {
                    switch (args.PropertyName)
                    {
                        case "Busy":
                        case "BusyMessage":
                            ManageWorkingState(_viewModel.Busy);

                            if (_viewModel.Busy && !string.IsNullOrEmpty(_viewModel.BusyMessage))
                            {
                                if (informationPanel == null)
                                    informationPanel =
                                        InformationPanel.GetInformationPanel(this, _viewModel.BusyMessage, 340, 150);
                                else
                                {
                                    informationPanel.Text = _viewModel.BusyMessage;
                                }
                            }
                            else if (informationPanel != null)
                            {
                                informationPanel.Dispose();
                                informationPanel = null;
                            }
                            break;
                        case "TemplateForUpload":
                            txtTemplatePath.Text = _viewModel.TemplateForUpload;
                            break;
                    }
                }, CancellationToken.None, TaskCreationOptions.None, uiSynchronizationContext);
            };

            tsbLoadEntities.Click += (sender, args) => _viewModel.LoadEntitiesCommand.Execute(sender);

            BindToObvervableCollection(cmbEntities.Items, _viewModel.Entities);
            cmbEntities.SelectedIndexChanged += (sender, args) =>
            {
                _viewModel.SelectedEntity = cmbEntities.SelectedItem as string;
            };

            BindToObvervableCollection(lvAttributes.Columns, _viewModel.ColumnHeadings, x => new ColumnHeader { Text = x });
            BindToObvervableCollection(lvAttributes.Items, _viewModel.Attributes);

            btnExport.Click += (sender, args) => _viewModel.DownloadTemplateCommand.Execute(sender);

            btnSelectTemplate.Click += (sender, args) => _viewModel.SelectTemplateForUploadCommand.Execute(sender);

            btnImport.Click += (sender, args) => _viewModel.UploadTemplateCommand.Execute(sender);

            cbCreate.CheckedChanged += (sender, args) => _viewModel.Create = cbCreate.Checked;
            cbUpdate.CheckedChanged += (sender, args) => _viewModel.Update = cbUpdate.Checked;
            cbDelete.CheckedChanged += (sender, args) => _viewModel.Delete = cbUpdate.Checked;

            tsbPublish.Click += (sender, args) => _viewModel.SaveChangesCommand.Execute(sender);
        }

        private void BindToObvervableCollection<T>(IList collection, ObservableCollection<T> observableCollection, Func<T, object> castItem = null)
        {
            var uiSynchronizationContext = TaskScheduler.FromCurrentSynchronizationContext();

            observableCollection.CollectionChanged += (sender, args) =>
            {
                Task.Factory.StartNew(() =>
                {
                    switch (args.Action)
                    {
                        case NotifyCollectionChangedAction.Add:
                            foreach (T item in args.NewItems)
                            {
                                object castedItem = castItem != null ? castItem(item) : item;

                                if (!collection.Contains(castedItem))
                                    collection.Add(castedItem);
                            }
                            break;
                        case NotifyCollectionChangedAction.Move:
                        case NotifyCollectionChangedAction.Remove:
                        case NotifyCollectionChangedAction.Replace:
                        case NotifyCollectionChangedAction.Reset:
                            collection.Clear();
                            foreach (T item in observableCollection)
                            {
                                object castedItem = castItem != null ? castItem(item) : item;

                                collection.Add(castedItem);
                            }
                            break;
                    }
                }, CancellationToken.None, TaskCreationOptions.None, uiSynchronizationContext);
            };
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
            _service = newService;
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

        #region Methods

        private void ManageWorkingState(bool working)
        {
            cmbEntities.Enabled = !working;
            gbSettings.Enabled = !working;
            gbAttributes.Enabled = !working;
            Cursor = working ? Cursors.WaitCursor : Cursors.Default;
        }

        #endregion Methods

    }
}