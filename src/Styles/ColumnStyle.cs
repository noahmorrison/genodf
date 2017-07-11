using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Genodf.Styles
{
    public class ColumnStyle : Base,
        IStyleable,
        ITableColumnProperties
    {
        public double? Width { get; set; }

        public void WriteStyle(XmlWriter xml)
        {
            if (!SetId())
                return;

            xml.WriteStartElement("style:style");
            xml.WriteAttributeString("style:family", "table-column");
            xml.WriteAttributeString("style:name", this.StyleId);

            this.WriteTableColumnProps(xml);

            WriteConditions(xml);
            xml.WriteEndElement();
        }
    }
}
