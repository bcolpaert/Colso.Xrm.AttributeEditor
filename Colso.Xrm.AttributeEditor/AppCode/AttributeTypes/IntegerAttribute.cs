using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class IntegerAttribute : AttributeMetadataBase<IntegerAttributeMetadata>
    {
        protected override void AddAdditionalMetadata(IntegerAttributeMetadata attribute)
        {
            attribute.Format = IntegerFormat.None;
            attribute.MaxValue = 100;
            attribute.MinValue = 0;
        }
    }
}
