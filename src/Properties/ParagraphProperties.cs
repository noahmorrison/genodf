using System.Xml;

public interface IParagraphProperties
{
    string TextAlign { get; set; }
}

internal static class ParagraphPropertiesExtension
{
    internal static void WriteParagraphProps(this IParagraphProperties props, XmlWriter xml)
    {
        if (!ParagraphIsStyled(props))
            return;

        xml.WriteStartElement("style:paragraph-properties");

        if (props.TextAlign != null)
            xml.WriteAttributeString("fo:text-align", props.TextAlign);

        xml.WriteEndElement();
    }

    internal static bool ParagraphIsStyled(this IParagraphProperties props)
    {
        return props.TextAlign != null;
    }
}
