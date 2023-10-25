using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class StatusAttribute : AttributeMetadataBase<StatusAttributeMetadata>
    {
        public string Options { get; set; }

        protected override void AddAdditionalMetadata(StatusAttributeMetadata attribute)
        {
            var options = ParseOptions();

            var optionCollection = new OptionMetadataCollection(options.Cast<OptionMetadata>().ToList());

            attribute.OptionSet = new OptionSetMetadata(optionCollection)
            {
                IsGlobal = false,
                OptionSetType = OptionSetType.Status,
            };
        }

        private List<StatusOptionMetadata> ParseOptions()
        {
            var options = Options.Split('\n').Select(x =>
                    x.Split(':'))
                .Select(x => new StatusOptionMetadata(int.Parse(x[0]), int.Parse(x[1]))
                {
                    Label = new Label(x[2], LanguageCode)
                }).ToList();
            return options;
        }

        public override void UpdateAttribute(IOrganizationService service)
        {
            var existingOptions = GetAttributeMetadata(service).OptionSet.Options;
            var updatedOptions = ParseOptions();

            foreach (var option in updatedOptions)
            {
                var existing = existingOptions.FirstOrDefault(x => x.Value == option.Value);

                if (existing == null)
                {
                    var request = new InsertStatusValueRequest
                    {
                        StateCode = option.State.Value,
                        Value = option.Value,
                        Label = option.Label,
                        AttributeLogicalName = LogicalName,
                        EntityLogicalName = Entity
                    };

                    service.Execute(request);
                }
            }

            base.UpdateAttribute(service);
        }

        private StatusAttributeMetadata GetAttributeMetadata(IOrganizationService service)
        {
            var request = new RetrieveMetadataChangesRequest
            {
                Query = new EntityQueryExpression
                {
                    AttributeQuery = new AttributeQueryExpression
                    {
                        Criteria =
                        {
                            Conditions =
                            {
                                new MetadataConditionExpression("logicalname", MetadataConditionOperator.Equals,
                                    LogicalName)
                            }
                        }
                    },
                    Criteria = new MetadataFilterExpression
                    {
                        Conditions =
                        {
                            new MetadataConditionExpression("logicalname", MetadataConditionOperator.Equals, Entity)
                        }
                    }
                }
            };

            var response = (RetrieveMetadataChangesResponse) service.Execute(request);

            return (StatusAttributeMetadata)response.EntityMetadata[0].Attributes.FirstOrDefault(x => x.LogicalName == LogicalName);
        }

        protected override void LoadAdditionalAttributeMetadata(StatusAttributeMetadata attribute)
        {
            var options = attribute.OptionSet.Options
                .OrderBy(x => x.Value)
                .Cast<StatusOptionMetadata>()
                .Select(x => $"{x.Value}:{x.State}:{x.Label?.UserLocalizedLabel?.Label}");

            Options = string.Join("\n", options);
        }
    }
}
