using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class MoneyAttribute : AttributeMetadataBase<MoneyAttributeMetadata>
    {
        public string Precision { get; set; }

        protected override void AddAdditionalMetadata(MoneyAttributeMetadata attribute)
        {
            attribute.MaxValue = 1000.00;
            attribute.MinValue = 0.00;
            attribute.Precision = int.Parse(Precision);
            attribute.PrecisionSource = 1;
            attribute.ImeMode = ImeMode.Disabled;
        }
    }
}
