using System.Xml;

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

    internal static bool TextIsStyled(this ITextProperties props)
    {
        return props.Fg != null
            || props.Bg != null
            || props.Bold;
    }
}
