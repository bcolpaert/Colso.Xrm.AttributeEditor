using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class DoubleAttribute : AttributeMetadataBase<DoubleAttributeMetadata>
    {
        public double? MaxValue { get; set; }
        public double? MinValue { get; set; }
        public int? Precision { get; set; }

        protected override void AddAdditionalMetadata(DoubleAttributeMetadata attribute)
        {
            attribute.MaxValue = MaxValue;
            attribute.MinValue = MinValue;
            attribute.Precision = Precision;
        }
        protected override void LoadAdditionalAttributeMetadata(DoubleAttributeMetadata attribute)
        {
            MaxValue = attribute.MaxValue;
            MinValue = attribute.MinValue;
            Precision = attribute.Precision;
        }
    }
}
