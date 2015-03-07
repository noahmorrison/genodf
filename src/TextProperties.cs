using System.Xml;

public interface ITextProperties
{
    string Bg { get; set; }
    string Fg { get; set; }
    bool Bold { get; set; }
}

public static class TextPropertiesExtension
{
    public static void WriteTextProps(this ITextProperties props, XmlWriter xml)
    {
        xml.WriteStartElement("style:text-properties");

        if (props.Fg != null)
            xml.WriteAttributeString("fo:color", props.Fg);
        if (props.Bg != null)
            xml.WriteAttributeString("fo:background-color", props.Bg);
        if (props.Bold)
            xml.WriteAttributeString("fo:font-weight", "bold");

        xml.WriteEndElement();
    }
}
