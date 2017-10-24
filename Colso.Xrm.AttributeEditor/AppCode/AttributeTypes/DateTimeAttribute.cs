using System;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class DateTimeAttribute : AttributeMetadataBase<DateTimeAttributeMetadata>
    {
        public string DateFormat { get; set; }

        protected override void AddAdditionalMetadata(DateTimeAttributeMetadata attribute)
        {
            DateTimeFormat dateFormat = DateTimeFormat.DateOnly;

            if (!string.IsNullOrWhiteSpace(DateFormat) && !DateTimeFormat.TryParse(DateFormat, out dateFormat))
                throw new Exception($"Could not parse DateFormat \"{DateFormat}\"");

            attribute.Format = dateFormat;
            attribute.ImeMode = ImeMode.Disabled;
        }
    }
}
