using System;
using System.Xml;

namespace Genodf
{
    public interface ITextProperties
    {
        string Bg { get; set; }
        string Fg { get; set; }
        bool Bold { get; set; }
    }

    internal static class TextPropertiesExtension
    {
        internal static void WriteTextProps(this ITextProperties props, XmlWriter xml)
        {
            if (!TextIsStyled(props))
                return;

            xml.WriteStartElement("style:text-properties");

            if (props.Fg != null)
                xml.WriteAttributeString("fo:color", props.Fg);
            if (props.Bg != null)
                xml.WriteAttributeString("fo:background-color", props.Bg);
            if (props.Bold)
                xml.WriteAttributeString("fo:font-weight", "bold");

            xml.WriteEndElement();
        }

        internal static void ReadTextProps(this ITextProperties self, XmlNode node)
        {
            if (node.Name != "style:text-properties")
                throw new ArgumentException("Xml node is not a text property node", "node");

            node.Attributes.IfHas("fo:color", c => self.Fg = c);
            node.Attributes.IfHas("fo:background-color", c => self.Bg = c);
            node.Attributes.IfHas("fo:fo:font-weight", _ => self.Bold = true);
        }

        internal static bool TextIsStyled(this ITextProperties props)
        {
            return props.Fg != null
                || props.Bg != null
                || props.Bold;
        }
    }
}