using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace Colso.Xrm.AttributeEditor.AppCode.AttributeTypes
{
    class PicklistAttribute : AttributeMetadataBase<PicklistAttributeMetadata>
    {
        public string Options { get; set; }

        public string GlobalOptionsetName { get; set; }

        protected override void AddAdditionalMetadata(PicklistAttributeMetadata attribute)
        {
            var options = Options.Split('\n').Select(x =>
                x.Split(':'))
                .Select(x => new OptionMetadata(new Label(x[1], LanguageCode), Int32.Parse(x[0]))).ToList();

            var optionCollection = new OptionMetadataCollection(options);

            var globalNames = !string.IsNullOrEmpty(GlobalOptionsetName) ? GlobalOptionsetName.Split(':') : null;

            attribute.OptionSet = new OptionSetMetadata(optionCollection)
            {
                IsGlobal = globalNames != null,
                Name = globalNames?[0],
                DisplayName = new Label(globalNames?[1], LanguageCode),
                OptionSetType = OptionSetType.Picklist,
            };
        }

        protected override void LoadAdditionalAttributeMetadata(PicklistAttributeMetadata attribute)
        {
            var options = attribute.OptionSet.Options.Select(x => $"{x.Value}:{x.Label}");

            Options = string.Join("\n", options);
        }
    }
}
