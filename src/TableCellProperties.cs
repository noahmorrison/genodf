using System.Xml;

public interface ITableCellProperties
{
    string Bg {get; set;}
    bool Border {get; set;}
    bool BorderTop {get; set;}
    bool BorderBottom {get; set;}
    bool BorderLeft {get; set;}
    bool BorderRight {get; set;}
}

public static class TableCellPropertiesExtension
{
    public static void WriteTableCellProps(this ITableCellProperties props, XmlWriter xml)
    {
        xml.WriteStartElement("style:table-cell-properties");

        if (props.Bg != null)
            xml.WriteAttributeString("fo:background-color", props.Bg);
        if (props.Border)
            xml.WriteAttributeString("fo:border", "00.6pt solid #000000");
        if (props.BorderTop)
            xml.WriteAttributeString("fo:border-top", "00.6pt solid #000000");
        if (props.BorderBottom)
            xml.WriteAttributeString("fo:border-bottom", "00.6pt solid #000000");
        if (props.BorderLeft)
            xml.WriteAttributeString("fo:border-left", "00.6pt solid #000000");
        if (props.BorderRight)
            xml.WriteAttributeString("fo:border-right", "00.6pt solid #000000");

        xml.WriteEndElement();
    }

    public static bool TableCellIsStyled(this ITableCellProperties props)
    {
        if (props.Bg != null)
            return true;
        if (props.Border)
            return true;
        if (props.BorderTop)
            return true;
        if (props.BorderBottom)
            return true;
        if (props.BorderLeft)
            return true;
        if (props.BorderRight)
            return true;
        return false;
    }
}