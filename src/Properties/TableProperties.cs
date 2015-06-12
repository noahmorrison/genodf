using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

public interface ITableProperties
{
    bool? Display { get; set; }
}

internal static class TablePropertiesExtension
{
    internal static void WriteTableProps(this ITableProperties props, XmlWriter xml)
    {
        if (!TableIsStyled(props))
            return;

        xml.WriteStartElement("style:table-properties");

        if (props.Display.HasValue)
            xml.WriteAttributeString("table:display", props.Display.ToString());

        xml.WriteEndElement();
    }

    internal static bool TableIsStyled(this ITableProperties props)
    {
        return props.Display.HasValue;
    }
}
