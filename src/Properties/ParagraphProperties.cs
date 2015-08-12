using System;
using System.Xml;

namespace Genodf
{
    public enum TextAlign
    {
        None,
        Left,
        Center,
        Right,
        Fill
    }

    public interface IParagraphProperties
    {
        TextAlign TextAlign { get; set; }
    }

    internal static class ParagraphPropertiesExtension
    {
        internal static void WriteParagraphProps(this IParagraphProperties props, XmlWriter xml)
        {
            if (!ParagraphIsStyled(props))
                return;

            xml.WriteStartElement("style:paragraph-properties");

            switch (props.TextAlign)
            {
                case TextAlign.Left:
                    xml.WriteAttributeString("fo:text-align", "start");
                    break;
                case TextAlign.Center:
                    xml.WriteAttributeString("fo:text-align", "center");
                    break;
                case TextAlign.Right:
                    xml.WriteAttributeString("fo:text-align", "end");
                    break;
                case TextAlign.Fill:
                    xml.WriteAttributeString("fo:text-align", "justify");
                    break;
                case TextAlign.None:
                    break;

                default:
                    throw new NotImplementedException("Text Align: " + props.TextAlign);
            }

            xml.WriteEndElement();
        }

        internal static void ReadParagraphProps(this IParagraphProperties self, XmlNode node)
        {
            if (node.Name != "style:paragraph-properties")
                throw new ArgumentException("Xml node is not a paragraph properties node", "node");

            node.Attributes.IfHas("fo:text-align", alignment =>
            {
                switch (alignment)
                {
                    case "start":
                        self.TextAlign = TextAlign.Left;
                        break;
                    case "center":
                        self.TextAlign = TextAlign.Center;
                        break;
                    case "end":
                        self.TextAlign = TextAlign.Right;
                        break;
                    case "justify":
                        self.TextAlign = TextAlign.Fill;
                        break;

                    default:
                        break;
                }
            });
        }

        internal static bool ParagraphIsStyled(this IParagraphProperties props)
        {
            return props.TextAlign != TextAlign.None;
        }
    }
}