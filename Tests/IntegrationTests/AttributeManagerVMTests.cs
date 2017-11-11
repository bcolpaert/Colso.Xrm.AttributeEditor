using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Threading;
using System.Threading.Tasks;
using Colso.Xrm.AttributeEditor;
using Microsoft.Xrm.Tooling.Connector;
using Moq;
using NUnit.Framework;

namespace Tests.IntegrationTests
{
    public class AttributeManagerVMTests
    {
        [SetUp]
        public void TestSetUp()
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
        }

        [Test]
        public void ProcessTest()
        {
            var client = new CrmServiceClient(ConfigurationManager.ConnectionStrings["CRM"].ConnectionString);
            var excelDocumentProcessor = new Mock<IExcelDocumentProcessor>();

            var uut = new AttributeManagerVM(() => client, excelDocumentProcessor.Object);

            // Load entities
            uut.LoadEntitiesCommand.Execute(null);

            SpinWait.SpinUntil(() => uut.Entities.Count > 0, TimeSpan.FromMinutes(1));

            // Select contact entity
            Assert.That(uut.Entities.Contains("contact"), Is.True);
            uut.SelectedEntity = "contact";

            SpinWait.SpinUntil(() => uut.Attributes.Count > 0, TimeSpan.FromMinutes(1));

            // Upload Spreadsheet
            uut.UploadTemplateCommand.Execute(null);

            // Verify that items are correct in attribute list
            // Save Entity
            // Verify Values
            // At beginning and end of test check for any attributes created during test and delete them
        }
    }
}
