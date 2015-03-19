using System.Xml;

public interface ITableColumnProperties
{
    double? Width { get; set; }
}

public static class TableColumnPropertiesExtension
{
    public static void WriteTableColumnProps(this ITableColumnProperties props, XmlWriter xml)
    {
        xml.WriteStartElement("style:table-column-properties");

        if (props.Width != null)
            xml.WriteAttributeString("style:column-width", props.Width.ToString() + "in");

        xml.WriteEndElement();
    }

    public static bool TableColumnIsStyled(this ITableColumnProperties props)
    {
        return props.Width != null;
    }
}
