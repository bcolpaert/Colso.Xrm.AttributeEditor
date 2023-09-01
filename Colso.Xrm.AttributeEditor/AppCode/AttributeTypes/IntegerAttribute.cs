using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class IntegerAttribute : AttributeMetadataBase<IntegerAttributeMetadata>
    {
        public string Format { get; set; }
        public int? MaxValue { get; set; }
        public int? MinValue { get; set; }

        protected override void AddAdditionalMetadata(IntegerAttributeMetadata attribute)
        {
            attribute.MaxValue = MaxValue;
            attribute.MinValue = MinValue;

            IntegerFormat format = IntegerFormat.None;
            if (!string.IsNullOrWhiteSpace(Format) && !IntegerFormat.TryParse(Format, out format))
                throw new ArgumentException($"Could not parse Format \"{Format}\"");
            attribute.Format = format;
        }

        protected override void LoadAdditionalAttributeMetadata(IntegerAttributeMetadata attribute)
        {
            MaxValue = attribute.MaxValue;
            MinValue = attribute.MinValue;
            Format = attribute.Format.ToString();
        }
    }
}
