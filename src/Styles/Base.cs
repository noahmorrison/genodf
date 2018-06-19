using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace Genodf.Styles
{
    public interface IStyleable
    {
        string StyleName { get; set; }
        string StyleId { get; }
        void WriteStyle(XmlWriter xml);

        bool AutomaticStyle { get; }
    }

    public class Base
    {
        private static Dictionary<string, string> _styles = new Dictionary<string, string>();
        private static Dictionary<string, int> _counts = new Dictionary<string, int>();

        private Dictionary<string, string> conditions = new Dictionary<string, string>();

        public virtual bool AutomaticStyle { get { return false; } }
        public string StyleName { get; set; }
        public string StyleId { get; private set; }

        public bool SetId()
        {
            if (!string.IsNullOrEmpty(StyleName))
            {
                StyleId = StyleName;
                return true;
            }

            var key = Serialize();

            if (_styles.ContainsKey(key))
            {
                StyleId = _styles[key];
                return false;
            }
            else
            {
                var styleType = this.GetType().Name;

                if (!_counts.ContainsKey(styleType))
                    _counts[styleType] = 0;

                StyleId = styleType + "-" + _counts[styleType].ToString();
                _styles[key] = StyleId;

                _counts[styleType]++;
                return true;
            }
        }

        internal static void Reset()
        {
            _styles = new Dictionary<string, string>();
            _counts = new Dictionary<string, int>();
        }

        private string Serialize()
        {
            Type type;
            switch (this.GetType().Name)
            {
                case "Cell":
                    type = (new CellStyle()).GetType();
                    break;

                case "Column":
                    type = (new ColumnStyle()).GetType();
                    break;

                case "Sheet":
                    type = (new TableStyle()).GetType();
                    break;

                default:
                    throw new NotImplementedException(this.GetType().Name);
            }

            var id = type.Name + " ";
            foreach (var propInfo in type.GetProperties(System.Reflection.BindingFlags.Public
                                                      | System.Reflection.BindingFlags.Instance))
            {
                if (propInfo.Name == "StyleId")
                    continue;

                id += propInfo.Name + "=" + propInfo.GetValue(this, null) + ", ";
            }

            foreach (var condition in conditions)
                id += " [" + condition.Key + ": " + condition.Value + "]";

            return id;
        }

        public void WriteConditions(XmlWriter xml)
        {
            foreach (var condition in conditions)
            {
                xml.WriteStartElement("style:map");
                xml.WriteAttributeString("style:condition", condition.Key);
                xml.WriteAttributeString("style:apply-style-name", condition.Value);
                xml.WriteEndElement();
            }
        }

        public void AddConditional(string condition, string globalStyleName)
        {
            if (conditions.ContainsKey(condition))
                if(conditions[condition] != globalStyleName)
                    throw new ArgumentException("Condition has already been added with another style");
                else
                    return;

            conditions.Add(condition, globalStyleName);
        }
    }
}
