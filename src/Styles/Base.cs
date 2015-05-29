using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Reflection;

namespace Genodf.Styles
{
    public interface IStyleable
    {
        string StyleId { get; }
        void WriteStyle(XmlWriter xml);
    }

    public class Base
    {
        private static Dictionary<string, string> _styles = new Dictionary<string, string>();
        private static Dictionary<string, int> _counts = new Dictionary<string, int>();

        public string StyleId { get; private set; }

        public bool SetId()
        {
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

            return id;
        }
    }
}
