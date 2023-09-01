using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class StringAttribute : AttributeMetadataBase<StringAttributeMetadata>
    {
        public int? MaxValue { get; set; }

        protected override void AddAdditionalMetadata(StringAttributeMetadata attribute)
        {
            attribute.MaxLength = MaxValue;
        }

        protected override void LoadAdditionalAttributeMetadata(StringAttributeMetadata attribute)
        {
            MaxValue = attribute.MaxLength;
        }
    }
}
