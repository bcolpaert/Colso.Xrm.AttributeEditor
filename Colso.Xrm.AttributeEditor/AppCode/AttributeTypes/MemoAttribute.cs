using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class MemoAttribute : AttributeMetadataBase<MemoAttributeMetadata>
    {
        public string Format { get; set; }
        public int? MaxLength { get; set; }

        protected override void AddAdditionalMetadata(MemoAttributeMetadata attribute)
        {
            StringFormat format = StringFormat.TextArea;
            if (!string.IsNullOrWhiteSpace(Format) && !StringFormat.TryParse(Format, out format))
                throw new ArgumentException($"Could not parse StringFormat \"{Format}\"");
            attribute.Format = format;

            attribute.ImeMode = ImeMode.Disabled;
            attribute.MaxLength = MaxLength;
        }

        protected override void LoadAdditionalAttributeMetadata(MemoAttributeMetadata attribute)
        {
            Format = attribute.Format.ToString();
            MaxLength = attribute.MaxLength;
        }
    }
}
