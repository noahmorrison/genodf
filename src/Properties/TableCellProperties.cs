using System;
using System.Xml;

namespace Genodf
{
    public interface ITableCellProperties
    {
        string Bg { get; set; }
        bool Border { get; set; }
        bool BorderTop { get; set; }
        bool BorderBottom { get; set; }
        bool BorderLeft { get; set; }
        bool BorderRight { get; set; }
    }

    internal static class TableCellPropertiesExtension
    {
        internal static void WriteTableCellProps(this ITableCellProperties props, XmlWriter xml)
        {
            if (!TableCellIsStyled(props))
                return;

            xml.WriteStartElement("style:table-cell-properties");

            if (props.Bg != null)
                xml.WriteAttributeString("fo:background-color", props.Bg);
            if (props.Border)
                xml.WriteAttributeString("fo:border", "0.3pt solid #000000");
            if (props.BorderTop)
                xml.WriteAttributeString("fo:border-top", "0.3pt solid #000000");
            if (props.BorderBottom)
                xml.WriteAttributeString("fo:border-bottom", "0.3pt solid #000000");
            if (props.BorderLeft)
                xml.WriteAttributeString("fo:border-left", "0.3pt solid #000000");
            if (props.BorderRight)
                xml.WriteAttributeString("fo:border-right", "0.3pt solid #000000");

            xml.WriteEndElement();
        }

        internal static void ReadTableCellProps(this ITableCellProperties self, XmlNode node)
        {
            if (node.Name != "style:table-cell-properties")
                throw new ArgumentException("Xml node is not a table cell property node", "node");

            node.Attributes.IfHas("fo:background-color", c => self.Bg = c);
            node.Attributes.IfHas("fo:border", _ => self.Border = true);
            node.Attributes.IfHas("fo:border-top", _ => self.BorderTop = true);
            node.Attributes.IfHas("fo:border-bottom", _ => self.BorderBottom = true);
            node.Attributes.IfHas("fo:border-left", _ => self.BorderLeft = true);
            node.Attributes.IfHas("fo:border-right", _ => self.BorderRight = true);
        }

        internal static bool TableCellIsStyled(this ITableCellProperties props)
        {
            return props.Bg != null
                || props.Border
                || props.BorderTop
                || props.BorderBottom
                || props.BorderLeft
                || props.BorderRight;
        }
    }
}