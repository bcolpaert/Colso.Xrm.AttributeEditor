using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Colso.Xrm.AttributeEditor.AppCode;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using XrmToolBox.Extensibility;

namespace Colso.Xrm.AttributeEditor
{
    class AttributeEditorViewModel
    {
        private readonly IMetadataHelper _metadataHelper;
        public IOrganizationService Service { get; set; }
        public List<EntityItem> Entities { get; set; }
        public bool WorkingState;

        public event EventHandler OnRequestConnection;
        public event EventHandler OnEntitiesListChanged;
        public event EventHandler OnWorkingStateChanged;
        public event EventHandler<GetInformationPanelEventArgs> OnGetInformationPanel;
        public event EventHandler<MessageBoxEventArgs> OnShowMessageBox;

        public AttributeEditorViewModel(IMetadataHelper metadataHelper)
        {
            _metadataHelper = metadataHelper;

            Entities = new List<EntityItem>();
        }

        public void LoadEntities()
        {
            if (Service == null)
            {
                OnRequestConnection?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                PopulateEntities();
            }
        }

        private void PopulateEntities()
        {
            if (!WorkingState)
            {
                // Reinit other controls
                Entities.Clear();
                OnEntitiesListChanged?.Invoke(this, EventArgs.Empty);
                WorkingState = true;
                OnWorkingStateChanged?.Invoke(this, EventArgs.Empty);

                var infoPanelEventArgs =
                    new GetInformationPanelEventArgs {Message = "Loading Entities...", Width = 340, Height = 150};
                OnGetInformationPanel?.Invoke(this, infoPanelEventArgs);
                var informationPanel = infoPanelEventArgs.Panel;

                // Launch treatment
                var bwFill = new BackgroundWorker();
                bwFill.DoWork += (sender, e) =>
                {
                    // Retrieve 
                    List<EntityMetadata> sourceList = _metadataHelper.RetrieveEntities(Service);

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
                        OnShowMessageBox?.Invoke(this, new MessageBoxEventArgs("An error occured: " + e.Error.Message, "Error", MessageBoxButtons.OK,
                            MessageBoxIcon.Error));
                    }
                    else
                    {
                        var items = (EntityItem[])e.Result;
                        if (items.Length == 0)
                        {
                            OnShowMessageBox?.Invoke(this, new MessageBoxEventArgs("The system does not contain any entities", "Warning", MessageBoxButtons.OK,
                                MessageBoxIcon.Warning));
                        }
                        else
                        {
                            Entities.AddRange(items);
                            OnEntitiesListChanged?.Invoke(this, EventArgs.Empty);
                        }
                    }

                    WorkingState = false;
                    OnWorkingStateChanged?.Invoke(this, EventArgs.Empty);
                };
                bwFill.RunWorkerAsync();
            }
        }

        public class GetInformationPanelEventArgs : EventArgs
        {
            public string Message { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
            public Panel Panel { get; set; }
        }

        public class MessageBoxEventArgs : EventArgs
        {
            public MessageBoxEventArgs() { }

            public MessageBoxEventArgs(string message, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
            {
                Message = message;
                Caption = caption;
                Buttons = buttons;
                Icon = icon;
            }

            public string Message { get; set; }
            public string Caption { get; set; }
            public MessageBoxButtons Buttons { get; set; }
            public MessageBoxIcon Icon { get; set; }
        }
    }
}
