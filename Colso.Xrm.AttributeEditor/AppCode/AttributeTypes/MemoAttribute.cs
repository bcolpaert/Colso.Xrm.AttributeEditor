using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class MemoAttribute : AttributeMetadataBase<MemoAttributeMetadata>
    {
        protected override void AddAdditionalMetadata(MemoAttributeMetadata attribute)
        {
            attribute.Format = StringFormat.TextArea;
            attribute.ImeMode = ImeMode.Disabled;
            attribute.MaxLength = 500;
        }
    }
}
