using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Genodf.Styles
{
    public class CellStyle : Base,
        IFormatable,
        ITableCellProperties,
        IParagraphProperties,
        ITextProperties
    {
        public IFormat Format { get; set; }
        public string Bg { get; set; }
        public string Fg { get; set; }
        public bool Bold { get; set; }
        public string TextAlign { get; set; }
        public bool Border { get; set; }
        public bool BorderTop { get; set; }
        public bool BorderBottom { get; set; }
        public bool BorderLeft { get; set; }
        public bool BorderRight { get; set; }

        public void WriteStyle(XmlWriter xml)
        {
            if (!SetId())
                return;

            if (Format != null)
                Format.WriteFormat(xml);

            xml.WriteStartElement("style:style");
            xml.WriteAttributeString("style:family", "table-cell");
            xml.WriteAttributeString("style:name", StyleId);

            if (Format != null)
                xml.WriteAttributeString("style:data-style-name", Format.Name);

            this.WriteTableCellProps(xml);
            this.WriteParagraphProps(xml);
            this.WriteTextProps(xml);

            xml.WriteEndElement();
        }
    }
}
