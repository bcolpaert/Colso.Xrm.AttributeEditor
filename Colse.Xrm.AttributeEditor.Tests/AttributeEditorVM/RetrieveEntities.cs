using System;
using System.Collections.Generic;
using System.Threading.Tasks;
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
        // Mocks
        private Mock<IMetadataHelper> _metadataHelperMock;
        private Mock<IOrganizationService> _serviceMock;
        private Mock<IDisposable> _infoPanelMock;

        private AttributeEditorViewModel _uut;
        private List<EntityMetadata> _entityMetadata;


        [SetUp]
        public void Setup()
        {
            _entityMetadata = new List<EntityMetadata>
            {
                new EntityMetadata(),
                new EntityMetadata()
            };

            _metadataHelperMock = new Mock<IMetadataHelper>();
            _metadataHelperMock.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(_entityMetadata);

            _serviceMock = new Mock<IOrganizationService>();

            _infoPanelMock = new Mock<IDisposable>();

            _uut = new AttributeEditorViewModel(_metadataHelperMock.Object);

            _uut.Service = new Mock<IOrganizationService>().Object;
            _uut.OnGetInformationPanel += (sender, args) => { args.Panel = _infoPanelMock.Object; };
        }

        [Test]
        public async Task RetrieveEntities_DoesNotRunWhenWorkingTrue()
        {
            // Arrange
            _uut.WorkingState = true;

            // Act
            await _uut.LoadEntities();

            // Assert
            _metadataHelperMock.Verify(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()), Times.Never);
        }

        [Test]
        public async Task RetrieveEntities_ClearsExistingEntities()
        {
            // Add an extra entity before loading entities
            _uut.Entities.Add(new EntityItem(new EntityMetadata()));

            // Act
            await _uut.LoadEntities();

            // Assert that extra entity is not includes in result
            Assert.That(_uut.Entities.Count, Is.EqualTo(_entityMetadata.Count));
        }

        [Test]
        public async Task RetrieveEntities_SetsWorkingStateToTrueBeforeLoadingEntities()
        {
            var entitiesLoaded = false;
            var workingState = false;

            _metadataHelperMock.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Callback(() => { entitiesLoaded = true; });

            _uut.OnWorkingStateChanged += (sender, args) =>
            {
                if (entitiesLoaded == false)
                    workingState = _uut.WorkingState;
            };

            await _uut.LoadEntities();

            Assert.That(workingState, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_SetsWorkingStateToFalseAfterLoadingEntities()
        {
            var retrieveEntitiesCalled = false;
            var workingState = true;

            _metadataHelperMock.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Callback(() => { retrieveEntitiesCalled = true; });

            _uut.OnWorkingStateChanged += (sender, args) =>
            {
                if (retrieveEntitiesCalled)
                    workingState = _uut.WorkingState;
            };

            await _uut.LoadEntities();

            Assert.That(workingState, Is.False);
        }

        [Test]
        public async Task RetrieveEntities_PopulatesEntityItems()
        {
            int count = 0;
            _uut.OnEntitiesListChanged += (sender, args) => { count = _uut.Entities.Count; };

            await _uut.LoadEntities();

            Assert.That(count, Is.EqualTo(_entityMetadata.Count));
        }

        [Test]
        public async Task RetrieveEntities_DisplaysInformationPanelBeforeLoadingEntities()
        {
            var retrieveEntitiesCalled = false;
            var infoPanelLoaded = false;

            _metadataHelperMock.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Callback(() => { retrieveEntitiesCalled = true; });

            _uut.OnGetInformationPanel += (sender, args) =>
            {
                if (!retrieveEntitiesCalled)
                    infoPanelLoaded = true;
            };

            await _uut.LoadEntities();

            Assert.That(infoPanelLoaded, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_DisposesInformationPanelAfterLoadingEntities()
        {
            var retrieveEntities = false;
            var infoPanelDisposed = false;

            _metadataHelperMock.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(_entityMetadata)
                .Callback(() => { retrieveEntities = true; });

            _infoPanelMock.Setup(x => x.Dispose()).Callback(() =>
            {
                if (!retrieveEntities)
                    throw new Exception("InfoPanel disposed before entities loaded");
                infoPanelDisposed = true;
            });

            await _uut.LoadEntities();

            Assert.That(infoPanelDisposed, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_ShowsMessageBoxOnError()
        {
            _metadataHelperMock.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Throws(new Exception());

            var messageBoxShown = false;
            _uut.OnShowMessageBox += (sender, args) => { messageBoxShown = true; };

            await _uut.LoadEntities();

            Assert.That(messageBoxShown, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_ShowsMessageWhenZeroEntitiesFound()
        {
            _metadataHelperMock.Setup(x => x.RetrieveEntities(It.IsAny<IOrganizationService>()))
                .Returns(new List<EntityMetadata>());

            var messageBoxShown = false;
            _uut.OnShowMessageBox += (sender, args) => { messageBoxShown = true; };

            await _uut.LoadEntities();

            Assert.That(messageBoxShown, Is.True);
        }

        [Test]
        public async Task RetrieveEntities_RequestsConnectionWhenConnectionNull()
        {
            _uut.Service = null;

            var connectionRequested = false;
            _uut.OnRequestConnection += (sender, args) => { connectionRequested = true; };

            await _uut.LoadEntities();

            Assert.That(connectionRequested, Is.True);
        }
    }
}