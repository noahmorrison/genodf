using System.Xml;

public interface ITableCellProperties
{
    string Bg {get; set;}
}

public static class TableCellPropertiesExtension
{
    public static void WriteTableCellProps(this ITableCellProperties props, XmlWriter xml)
    {
        xml.WriteStartElement("style:table-cell-properties");

        if (props.Bg != null)
            xml.WriteAttributeString("fo:background-color", props.Bg);

        xml.WriteEndElement();
    }
}
