using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class MoneyAttribute : AttributeMetadataBase<MoneyAttributeMetadata>
    {
        public double? MaxValue { get; set; }
        public double? MinValue { get; set; }
        public int? Precision { get; set; }

        protected override void AddAdditionalMetadata(MoneyAttributeMetadata attribute)
        {
            attribute.MaxValue = MaxValue;
            attribute.MinValue = MinValue;

            attribute.PrecisionSource = Precision;
            attribute.ImeMode = ImeMode.Disabled;
        }

        protected override void LoadAdditionalAttributeMetadata(MoneyAttributeMetadata attribute)
        {
            MaxValue = attribute.MaxValue;
            MinValue = attribute.MinValue;
            Precision = attribute.PrecisionSource;
        }
    }
}
