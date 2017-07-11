using Genodf.Styles;

namespace Genodf
{
    public class Column : ColumnStyle
    {
        public void Write(XmlWriter xml)
        {
            xml.WriteStartElement("table:table-column");

            xml.WriteAttributeString("table:style-name", StyleId);

            xml.WriteEndElement();
        }
    }
}
