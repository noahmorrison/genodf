using System.Xml;

namespace Genodf
{
    public class Column : ITableColumnProperties
    {
        private int index;

        public double? Width { get; set; }

        public Column(int index)
        {
            this.index = index;
        }

        public bool IsStyled()
        {
            return this.TableColumnIsStyled();
        }

        public void WriteStyle(XmlWriter xml)
        {
            if (!this.IsStyled())
                return;

            xml.WriteStartElement("style:style");
            xml.WriteAttributeString("style:family", "table-column");
            xml.WriteAttributeString("style:name", "co" + this.index);

            this.WriteTableColumnProps(xml);

            xml.WriteEndElement();
        }

        public void Write(XmlWriter xml)
        {
            xml.WriteStartElement("table:table-column");

            if (this.IsStyled())
                xml.WriteAttributeString("table:style-name", "co" + this.index);

            xml.WriteEndElement();
        }
    }
}