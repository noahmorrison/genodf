using System;
using System.Xml;

namespace Genodf
{
    public interface ITableProperties
    {
        bool? Display { get; set; }
    }

    internal static class TablePropertiesExtension
    {
        internal static void WriteTableProps(this ITableProperties props, XmlWriter xml)
        {
            if (!TableIsStyled(props))
                return;

            xml.WriteStartElement("style:table-properties");

            if (props.Display.HasValue)
                xml.WriteAttributeString("table:display", props.Display.ToString());

            xml.WriteEndElement();
        }

        internal static void ReadTableProps(this ITableProperties self, XmlNode node)
        {
            if (node.Name != "style:table-properties")
                throw new ArgumentException("Xml node is not a table property node", "node");

            node.Attributes.IfHas("table:display", value => self.Display = bool.Parse(value));
        }

        internal static bool TableIsStyled(this ITableProperties props)
        {
            return props.Display.HasValue;
        }
    }
}