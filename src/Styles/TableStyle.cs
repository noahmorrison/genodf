using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Genodf.Styles
{
    public class TableStyle : Base,
        IStyleable,
        ITableProperties
    {
        public bool? Display { get; set; }
        public string MasterPageStyle { get; set; }

        public TableStyle()
        {
            MasterPageStyle = "Default";
        }

        public void WriteStyle(XmlWriter xml)
        {
            if (!SetId())
                return;

            xml.WriteStartElement("style:style");
            xml.WriteAttributeString("style:family", "table");
            xml.WriteAttributeString("style:name", this.StyleId);
            xml.WriteAttributeString("style:master-page-name", MasterPageStyle);

            this.WriteTableProps(xml);

            xml.WriteEndElement();
        }
    }
}
