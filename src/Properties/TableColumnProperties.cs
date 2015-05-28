using System.Xml;

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

    internal static bool TableColumnIsStyled(this ITableColumnProperties props)
    {
        return props.Width != null;
    }
}
