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
    public class EntityItem : IEquatable<EntityItem>, IComparable<EntityItem>
    {
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

        public bool Equals(EntityItem other)
        {
            if (record == null) return false;
            if (other == null) return false;

            return record.LogicalName.Equals(other.LogicalName);
        }

        public int CompareTo(EntityItem other)
        {
            // A null value means that this object is greater.
            if (other == null)
                return 1;

            if (record == null)
                return -1;

            
            return this.DisplayName.CompareTo(other.DisplayName);
        }
    }
}