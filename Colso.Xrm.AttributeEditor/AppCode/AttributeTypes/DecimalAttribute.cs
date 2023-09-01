using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class DecimalAttribute : AttributeMetadataBase<DecimalAttributeMetadata>
    {
        public decimal? MaxValue { get; set; }
        public decimal? MinValue { get; set; }
        public int? Precision { get; set; }

        protected override void AddAdditionalMetadata(DecimalAttributeMetadata attribute)
        {
            attribute.MaxValue = MaxValue;
            attribute.MinValue = MinValue;
            attribute.Precision = Precision;
        }

        protected override void LoadAdditionalAttributeMetadata(DecimalAttributeMetadata attribute)
        {
            MaxValue = attribute.MaxValue;
            MinValue = attribute.MinValue;
            Precision = attribute.Precision;
        }
    }
}
