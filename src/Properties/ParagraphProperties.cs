using System.Xml;

public interface IParagraphProperties
{
    string TextAlign { get; set; }
}

public static class ParagraphPropertiesExtension
{
    public static void WriteParagraphProps(this IParagraphProperties props, XmlWriter xml)
    {
        xml.WriteStartElement("style:paragraph-properties");

        if (props.TextAlign != null)
            xml.WriteAttributeString("fo:text-align", props.TextAlign);

        xml.WriteEndElement();
    }

    public static bool ParagraphIsStyled(this IParagraphProperties props)
    {
        return props.TextAlign != null;
    }
}
