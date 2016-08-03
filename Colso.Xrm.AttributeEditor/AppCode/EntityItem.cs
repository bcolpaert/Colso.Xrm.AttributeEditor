using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Colso.Xrm.AttributeEditor.AppCode
{
    public class EntityItem
    {
        private readonly IOrganizationService sourceService;
        private readonly EntityMetadata record;

        public string DisplayName
        {
            get
            {
                return record.DisplayName == null ? 
                    string.Empty : 
                    record.DisplayName.UserLocalizedLabel == null ? 
                    record.DisplayName.LocalizedLabels.Select(l => l.Label).FirstOrDefault() : 
                    record.DisplayName.UserLocalizedLabel.Label;
            }
        }

        public string LogicalName
        {
            get
            {
                return record.LogicalName;
            }
        }
        public int LanguageCode
        {
            get
            {
                return record.DisplayName == null || record.DisplayName.UserLocalizedLabel  == null ?
                    1033 :
                    record.DisplayName.UserLocalizedLabel.LanguageCode;
            }
        }

        public EntityItem(EntityMetadata record)
        {
            this.record = record;
        }
    }
}