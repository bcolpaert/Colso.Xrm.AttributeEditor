using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class StringAttribute : AttributeMetadataBase<StringAttributeMetadata>
    {
        protected override void AddAdditionalMetadata(StringAttributeMetadata attribute)
        {
            attribute.MaxLength = 100;
        }
    }
}
