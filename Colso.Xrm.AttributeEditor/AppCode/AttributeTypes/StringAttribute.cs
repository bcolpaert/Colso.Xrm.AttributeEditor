using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class StringAttribute : AttributeMetadataBase<StringAttributeMetadata>
    {
        public int? MaxLength { get; set; }

        protected override void AddAdditionalMetadata(StringAttributeMetadata attribute)
        {
            attribute.MaxLength = MaxLength;
        }

        protected override void LoadAdditionalAttributeMetadata(StringAttributeMetadata attribute)
        {
            MaxLength = attribute.MaxLength;
        }
    }
}
