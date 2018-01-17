using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Castle.Components.DictionaryAdapter;
using Colso.Xrm.AttributeEditor;
using Colso.Xrm.AttributeEditor.AppCode;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Moq;
using NUnit.Framework;

namespace Colse.Xrm.AttributeEditor.Tests.AttributeEditorVM
{
    public class RetrieveEntities
    {
        [Test]
        public async Task RetrieveEntities_DoesNotRunWhenWorkingTrue()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            var uut = new AttributeEditorViewModel(metadataHelper.Object);
            uut.WorkingState = true;

            await uut.LoadEntities();

            metadataHelper.Verify(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()), Times.Never);
        }

        [Test]
        public async Task RetrieveEntities_ClearsExistingEntities()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(new List<EntityMetadata>());
            var uut = new AttributeEditorViewModel(metadataHelper.Object);

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            uut.Entities.Add(new EntityItem(new EntityMetadata()));

            await uut.LoadEntities();

            Assert.That(uut.Entities.Count == 0);
        }

        [Test]
        public async Task RetrieveEntities_SetsWorkingStateToTrueBeforeLoadingEntities()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            var uut = new AttributeEditorViewModel(metadataHelper.Object);
            var metadataLoaded = false;
            var workingState = false;

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(new List<EntityMetadata>())
                .Callback(() => { metadataLoaded = true; });
            uut.OnWorkingStateChanged += (sender, args) =>
            {
                if (metadataLoaded == false)
                    workingState = uut.WorkingState;
            };

            await uut.LoadEntities();

            Assert.That(workingState, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_SetsWorkingStateToFalseAfterLoadingEntities()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            var uut = new AttributeEditorViewModel(metadataHelper.Object);

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            var metadataLoaded = false;
            var workingState = true;

            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(new List<EntityMetadata>())
                .Callback(() => { metadataLoaded = true; });

            uut.OnWorkingStateChanged += (sender, args) =>
            {
                if (metadataLoaded)
                    workingState = uut.WorkingState;
            };

            await uut.LoadEntities();

            Assert.That(workingState, Is.False);
        }

        [Test]
        public async Task RetrieveEntities_PopulatesEntityItems()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            var uut = new AttributeEditorViewModel(metadataHelper.Object);

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            var metadata = new List<EntityMetadata>
            {
                new EntityMetadata(),
                new EntityMetadata()
            };
            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>())).Returns(metadata);

            int count = 0;
            uut.OnEntitiesListChanged += (sender, args) => { count = uut.Entities.Count; };

            await uut.LoadEntities();

            Assert.That(count, Is.EqualTo(2));
        }

        [Test]
        public async Task RetrieveEntities_DisplaysInformationPanelBeforeLoadingEntities()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            var uut = new AttributeEditorViewModel(metadataHelper.Object);

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            var entitiesLoaded = false;
            var infoPanelLoaded = false;

            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(new List<EntityMetadata>()).Callback(
                    () => { entitiesLoaded = true; });

            uut.OnGetInformationPanel += (sender, args) =>
            {
                if (!entitiesLoaded)
                    infoPanelLoaded = true;
            };

            await uut.LoadEntities();

            Assert.That(infoPanelLoaded, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_DisposesInformationPanelAfterLoadingEntities()
        {
            var metadataHelper = new Mock<IMetadataHelper>();
            

            var uut = new AttributeEditorViewModel(metadataHelper.Object);

            var entitiesLoaded = false;
            var infoPanelDisposed = false;

            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(new List<EntityMetadata>())
                .Callback(() => { entitiesLoaded = true; });

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            infoPanel.Setup(x => x.Dispose()).Callback(() =>
            {
                if (!entitiesLoaded)
                    throw new Exception("InfoPanel disposed before entities loaded");
                infoPanelDisposed = true;
            });

            await uut.LoadEntities();

            Assert.That(infoPanelDisposed, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_ShowsMessageBoxOnError()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            var uut = new AttributeEditorViewModel(metadataHelper.Object);

            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Throws(new Exception());

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            var messageBoxShown = false;
            uut.OnShowMessageBox += (sender, args) => { messageBoxShown = true; };

            await uut.LoadEntities();

            Assert.That(messageBoxShown, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_ShowsMessageWhenZeroEntitiesFound()
        {
            var metadataHelper = new Mock<IMetadataHelper>();

            var uut = new AttributeEditorViewModel(metadataHelper.Object);

            metadataHelper.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(new List<EntityMetadata>());

            var infoPanel = new Mock<IDisposable>();
            uut.OnGetInformationPanel += (sender, args) => { args.Panel = infoPanel.Object; };

            var messageBoxShown = false;
            uut.OnShowMessageBox += (sender, args) => { messageBoxShown = true; };

            await uut.LoadEntities();

            Assert.That(messageBoxShown, Is.True);
        }
    }
}