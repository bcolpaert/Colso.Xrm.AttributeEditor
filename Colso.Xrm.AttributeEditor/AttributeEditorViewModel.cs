using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Colso.Xrm.AttributeEditor.AppCode;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor
{
    public class AttributeEditorViewModel
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

        public async Task LoadEntities()
        {
            if (Service == null)
            {
                OnRequestConnection?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                await PopulateEntities();
            }
        }

        private async Task PopulateEntities()
        {
            if (WorkingState)
                return;

            // Reinit other controls
            Entities.Clear();
            WorkingState = true;
            OnWorkingStateChanged?.Invoke(this, EventArgs.Empty);

            var infoPanelEventArgs =
                new GetInformationPanelEventArgs { Message = "Loading Entities...", Width = 340, Height = 150 };
            OnGetInformationPanel?.Invoke(this, infoPanelEventArgs);
            var informationPanel = infoPanelEventArgs.Panel;

            var sourceEntitiesList = new List<EntityItem>();

            try
            {
                await Task.Factory.StartNew(() =>
                {
                    var sourceList = _metadataHelper.RetrieveEntities(Service);

                    foreach (EntityMetadata entity in sourceList)
                        sourceEntitiesList.Add(new EntityItem(entity));
                });

                informationPanel.Dispose();

                if (sourceEntitiesList.Count == 0)
                {
                    OnShowMessageBox?.Invoke(this, new MessageBoxEventArgs("The system does not contain any entities", "Warning", MessageBoxButtons.OK,
                                        MessageBoxIcon.Warning));
                }
                else
                {
                    Entities.AddRange(sourceEntitiesList);
                    OnEntitiesListChanged?.Invoke(this, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                OnShowMessageBox?.Invoke(this, new MessageBoxEventArgs("An error occured: " + ex.Message, "Error", MessageBoxButtons.OK,
                                MessageBoxIcon.Error));
            }

            WorkingState = false;
            OnWorkingStateChanged?.Invoke(this, EventArgs.Empty);
        }

        public class GetInformationPanelEventArgs : EventArgs
        {
            public string Message { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
            public IDisposable Panel { get; set; }
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
