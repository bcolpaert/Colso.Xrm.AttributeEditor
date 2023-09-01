using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Collections.Generic;
using System.Linq;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class BooleanAttribute : AttributeMetadataBase<BooleanAttributeMetadata>
    {
        public string Options { get; set; }

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
            attribute.OptionSet = ParseOptions();
        }

        private BooleanOptionSetMetadata ParseOptions()
        {
            var options = Options.Split('\n').Select(x => x.Split(':'))
                    .Where(x => x != null && x.Length > 1)
                    .Select(x => new OptionMetadata(new Label(x[1], LanguageCode), int.Parse(x[0])));

            var trueOption = options.Where(x => x.Value == 1).FirstOrDefault();
            var falseOption = options.Where(x => x.Value == 0).FirstOrDefault();

            if (trueOption == null) trueOption = new OptionMetadata(new Label("True", LanguageCode), 1);
            if (falseOption == null) falseOption = new OptionMetadata(new Label("False", LanguageCode), 0);

            return new BooleanOptionSetMetadata(trueOption, falseOption);
        }

        protected override void LoadAdditionalAttributeMetadata(BooleanAttributeMetadata attribute)
        {
            var options = new List<string>();
            options.Add($"{attribute.OptionSet.TrueOption.Value}:{attribute.OptionSet.TrueOption.Label?.UserLocalizedLabel?.Label}");
            options.Add($"{attribute.OptionSet.FalseOption.Value}:{attribute.OptionSet.FalseOption.Label?.UserLocalizedLabel?.Label}");

            Options = string.Join("\n", options);
        }
    }
}
