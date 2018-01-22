using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class DoubleAttribute : AttributeMetadataBase<DoubleAttributeMetadata>
    {
        protected override void AddAdditionalMetadata(DoubleAttributeMetadata attribute)
        {
            attribute.MaxValue = 100;
            attribute.MinValue = 0;
            attribute.Precision = 1;
        }
    }
}
