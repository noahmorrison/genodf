using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Genodf.Styles
{
    public class CellStyle : Base,
        IStyleable,
        IFormatable,
        ITableCellProperties,
        IParagraphProperties,
        ITextProperties
    {
        public IFormat Format { get; set; }
        public string Bg { get; set; }
        public string Fg { get; set; }
        public bool Bold { get; set; }
        public TextAlign TextAlign { get; set; }
        public bool Border { get; set; }
        public bool BorderTop { get; set; }
        public bool BorderBottom { get; set; }
        public bool BorderLeft { get; set; }
        public bool BorderRight { get; set; }

        public CellStyle()
        { }

        public CellStyle(XmlNode node)
        {
            if (node.Attributes["style:name"] != null)
                StyleName = node.Attributes["style:name"].Value;

            foreach (XmlNode props in node.ChildNodes)
                switch (props.Name)
                {
                    case "style:paragraph-properties":
                        this.ReadParagraphProps(props);
                        break;
                    case "style:table-cell-properties":
                        this.ReadTableCellProps(props);
                        break;
                    case "style:text-properties":
                        this.ReadTextProps(props);
                        break;
                    case "style:map":
                        var cond = props.Attributes["style:condition"].Value;
                        var style = props.Attributes["style:apply-style-name"].Value;

                        this.AddConditional(cond, style);
                        break;

                    default:
                        throw new ArgumentException("Invalid properties type");
                }
        }

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
                xml.WriteAttributeString("style:data-style-name", Format.FormatId);

            this.WriteTableCellProps(xml);
            this.WriteParagraphProps(xml);
            this.WriteTextProps(xml);

            WriteConditions(xml);
            xml.WriteEndElement();
        }
    }
}
