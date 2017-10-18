using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Colso.Xrm.AttributeEditor.AppCode
{
    public class EntityHelper
    { 

        private string entity;
        public readonly int languageCode = 1033;
        private readonly IOrganizationService service;

        public EntityHelper(string entitylogicalname, int languageCode, IOrganizationService service)
        {
            this.entity = entitylogicalname;
            this.languageCode = languageCode;
            this.service = service;
        }

        public void CreateAttribute(string logicalname, string displayname, string type, string requirement)
        {
            var requiredlevel = AttributeRequiredLevel.None;
            Enum.TryParse(requirement, out requiredlevel);

            var attribute = GetAttributeMetadata(type, logicalname, displayname, requiredlevel);

            var request = new CreateAttributeRequest
            {
                EntityName = entity,
                Attribute = attribute
            };

            service.Execute(request);
        }

        public void UpdateAttribute(string logicalname, string displayname, string requirement)
        {
            var requiredlevel = AttributeRequiredLevel.None;
            Enum.TryParse(requirement, out requiredlevel);

            var attribute = new AttributeMetadata();
            // Set base properties
            attribute.LogicalName = logicalname;
            attribute.DisplayName = new Label(displayname, languageCode);
            attribute.RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredlevel);

            var request = new UpdateAttributeRequest
            {
                EntityName = entity,
                Attribute = attribute
            };

            service.Execute(request);
        }

        public void DeleteAttribute(string logicalname)
        {
            var request = new DeleteAttributeRequest
            {
                EntityLogicalName = entity,
                LogicalName = logicalname
            };

            service.Execute(request);
        }

        public void Publish()
        {
            var pubRequest = new PublishXmlRequest();
            pubRequest.ParameterXml = string.Format(@"<importexportxml>
                                                           <entities>
                                                              <entity>{0}</entity>
                                                           </entities>
                                                        </importexportxml>",
                                                    entity);

            service.Execute(pubRequest);
        }

        private AttributeMetadata GetAttributeMetadata(string type, string logicalname, string displayname, AttributeRequiredLevel requiredlevel)
        {
            AttributeMetadata attribute = null;

            if (type != null)
            {
                switch (type.ToLower())
                {
                    case "boolean":
                        attribute = new BooleanAttributeMetadata
                        {
                            // Set extended properties
                            OptionSet = new BooleanOptionSetMetadata(
                            new OptionMetadata(new Label("True", languageCode), 1),
                            new OptionMetadata(new Label("False", languageCode), 0)
                            )
                        };
                        break;
                    case "datetime":
                        // Create a date time attribute
                        attribute = new DateTimeAttributeMetadata
                        {
                            // Set extended properties
                            Format = DateTimeFormat.DateOnly,
                            ImeMode = ImeMode.Disabled
                        };
                        break;
                    case "double":
                        // Create a decimal attribute	
                        attribute = new DoubleAttributeMetadata
                        {
                            // Set extended properties
                            MaxValue = 100,
                            MinValue = 0,
                            Precision = 1
                        };
                        break;
                    case "decimal":
                        // Create a decimal attribute	
                        attribute = new DecimalAttributeMetadata
                        {
                            // Set extended properties
                            MaxValue = 100,
                            MinValue = 0,
                            Precision = 1
                        };
                        break;
                    case "integer":
                        // Create a integer attribute	
                        attribute = new IntegerAttributeMetadata
                        {
                            // Set extended properties
                            Format = IntegerFormat.None,
                            MaxValue = 100,
                            MinValue = 0
                        };
                        break;
                    case "memo":
                        // Create a memo attribute 
                        attribute = new MemoAttributeMetadata
                        {
                            // Set extended properties
                            Format = StringFormat.TextArea,
                            ImeMode = ImeMode.Disabled,
                            MaxLength = 500
                        };
                        break;
                    case "money":
                        // Create a money attribute	
                        MoneyAttributeMetadata moneyAttribute = new MoneyAttributeMetadata
                        {
                            // Set extended properties
                            MaxValue = 1000.00,
                            MinValue = 0.00,
                            Precision = 1,
                            PrecisionSource = 1,
                            ImeMode = ImeMode.Disabled
                        };
                        break;
                    case "picklist":
                        // Create a picklist attribute	
                        attribute = new PicklistAttributeMetadata
                        {
                            // Set extended properties
                            // Build local picklist options
                            OptionSet = new OptionSetMetadata
                            {
                                IsGlobal = false,
                                OptionSetType = OptionSetType.Picklist
                            }
                        };
                        break;
                    case "string":
                        // Create a string attribute
                        attribute = new StringAttributeMetadata
                        {
                            // Set extended properties
                            MaxLength = 100
                        };
                        break;
                    default:
                        throw new ArgumentException(string.Format("Unexpected attribute type: {0}", type));
                }

                // Set base properties
                attribute.SchemaName = logicalname;
                attribute.DisplayName = new Label(displayname, languageCode);
                attribute.RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredlevel);
            }

            return attribute;
        }
    }
}
