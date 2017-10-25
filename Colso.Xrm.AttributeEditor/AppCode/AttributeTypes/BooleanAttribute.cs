using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class BooleanAttribute : AttributeMetadataBase<BooleanAttributeMetadata>
    {
        public override void CreateAttribute(IOrganizationService service)
        {
            var attribute = GetAttributeMetadata();

            var request = new CreateAttributeRequest
            {
                EntityName = Entity,
                Attribute = attribute
            };

            service.Execute(request);
        }

        public override void UpdateAttribute(IOrganizationService service)
        {
            var attribute = GetAttributeMetadata();

            var request = new UpdateAttributeRequest
            {
                EntityName = Entity,
                Attribute = attribute
            };

            service.Execute(request);
        }

        protected override void AddAdditionalMetadata(BooleanAttributeMetadata attribute)
        {
            attribute.OptionSet = new BooleanOptionSetMetadata(
                new OptionMetadata(new Label("True", LanguageCode), 1),
                new OptionMetadata(new Label("False", LanguageCode), 0)
            );
        }
    }
}
