using System;
using System.Xml;

namespace Genodf
{
    public interface ITableColumnProperties
    {
        double? Width { get; set; }
    }

    internal static class TableColumnPropertiesExtension
    {
        internal static void WriteTableColumnProps(this ITableColumnProperties props, XmlWriter xml)
        {
            if (!props.TableColumnIsStyled())
                return;

            xml.WriteStartElement("style:table-column-properties");

            if (props.Width != null)
                xml.WriteAttributeString("style:column-width", props.Width.ToString() + "in");

            xml.WriteEndElement();
        }

        internal static void ReadTableColumnProps(this ITableColumnProperties self, XmlNode node)
        {
            if (node.Name != "style:table-column-properties")
                throw new ArgumentException("Xml node is not a table column property node", "node");

            node.Attributes.IfHas("style:column-width", value =>
                self.Width = double.Parse(value.Replace("in", string.Empty)));
        }

        internal static bool TableColumnIsStyled(this ITableColumnProperties props)
        {
            return props.Width != null;
        }
    }
}
