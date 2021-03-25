using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    abstract class AttributeMetadataBase<T> : AttributeBase where T : AttributeMetadata, new()
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

        protected T GetAttributeMetadata()
        {
            var attribute = new T
            {
                LogicalName = LogicalName,
                SchemaName = SchemaName,
                RequiredLevel = ParseRequiredLevel(Requirement)
            };

            // Only add labels if we have a value
            if (!string.IsNullOrEmpty(DisplayName)) attribute.DisplayName = new Label(DisplayName, LanguageCode);
            if (!string.IsNullOrEmpty(Description)) attribute.Description = new Label(Description, LanguageCode);


            AddAdditionalMetadata(attribute);

            return attribute;
        }

        protected override void LoadAdditionalAttributeMetadata(AttributeMetadata attribute)
        {
            LoadAdditionalAttributeMetadata((T)attribute);
        }

        protected virtual void LoadAdditionalAttributeMetadata(T attribute) { }

        protected abstract void AddAdditionalMetadata(T attribute);
    }
}
