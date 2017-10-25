using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class LookupAttribute : AttributeMetadataBase<LookupAttributeMetadata>
    {
        public string LookupTarget { get; set; }

        public override void CreateAttribute(IOrganizationService service)
        {
            var request = new CreateOneToManyRequest
            {
                OneToManyRelationship = GetRelationshipMetadata(),
                Lookup = GetAttributeMetadata()
            };

            service.Execute(request);
        }

        public override void UpdateAttribute(IOrganizationService service)
        {
            var relationshipRequest = new UpdateRelationshipRequest
            {
                Relationship = GetRelationshipMetadata()
            };

            service.Execute(relationshipRequest);

            var lookupRequest = new UpdateAttributeRequest
            {
                EntityName = Entity,
                Attribute = GetAttributeMetadata()
            };

            service.Execute(lookupRequest);
        }

        protected override void AddAdditionalMetadata(LookupAttributeMetadata attribute)
        {
            LookupTarget = LookupTarget;
        }

        private OneToManyRelationshipMetadata GetRelationshipMetadata()
        {
            var prefix = LogicalName.Split('_')[0];

            return new OneToManyRelationshipMetadata
            {
                ReferencedEntity = LookupTarget,
                ReferencingEntity = Entity,
                SchemaName = $"{prefix}_{LookupTarget}_{Entity}",
                AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                {
                    Behavior = AssociatedMenuBehavior.UseLabel,
                    Group = AssociatedMenuGroup.Details,
                    Label = new Label(LookupTarget, LanguageCode),
                    Order = 10000
                },
                CascadeConfiguration = new CascadeConfiguration
                {
                    Assign = CascadeType.NoCascade,
                    Delete = CascadeType.RemoveLink,
                    Merge = CascadeType.NoCascade,
                    Reparent = CascadeType.NoCascade,
                    Share = CascadeType.NoCascade,
                    Unshare = CascadeType.NoCascade
                }
            };
        }

        protected override void LoadAdditionalAttributeMetadata(LookupAttributeMetadata attribute)
        {
            LookupTarget = attribute.Targets[0];
        }
    }
}
