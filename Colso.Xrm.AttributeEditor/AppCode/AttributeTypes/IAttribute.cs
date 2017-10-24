using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    public interface IAttribute
    {
        string DisplayName { get; set; }
        string Entity { get; set; }
        string LogicalName { get; set; }
        string Requirement { get; set; }

        void CreateAttribute(IOrganizationService service);
        void DeleteAttribute(IOrganizationService service);
        void UpdateAttribute(IOrganizationService service);
        void LoadFromAttributeMetadata(AttributeMetadata attribute);


        void LoadFromAttributeMetadataRow(AttributeMetadataRow row);
        AttributeMetadataRow ToAttributeMetadataRow();
    }
}