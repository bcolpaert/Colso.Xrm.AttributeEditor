using System;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class DateTimeAttribute : AttributeMetadataBase<DateTimeAttributeMetadata>
    {
        public string Format { get; set; }

        protected override void AddAdditionalMetadata(DateTimeAttributeMetadata attribute)
        {
            DateTimeFormat dateFormat = DateTimeFormat.DateOnly;
            if (!string.IsNullOrWhiteSpace(Format) && !DateTimeFormat.TryParse(Format, out dateFormat))
                throw new ArgumentException($"Could not parse DateFormat \"{Format}\"");
            attribute.Format = dateFormat;

            attribute.ImeMode = ImeMode.Disabled;
        }

        protected override void LoadAdditionalAttributeMetadata(DateTimeAttributeMetadata attribute)
        {
            Format = attribute.Format.ToString();
        }
    }
}
