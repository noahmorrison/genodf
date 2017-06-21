using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Genodf.Styles
{
    public class PageLayoutStyle : Base,
        IStyleable,
        IPageLayoutProperties
    {
        public string PageWidth { get; set; }
        public string PageHeight { get; set; }
        public int? FitInPages { get; set; }

        public override bool AutomaticStyle { get { return true; } }

        public PageLayoutStyle(string name)
        {
            StyleName = name;
        }

        public PageLayoutStyle(XmlNode node)
        {
            if (node.Attributes["style:name"] != null)
                StyleName = node.Attributes["style:name"].Value;

            foreach (XmlNode props in node.ChildNodes)
                switch (props.Name)
                {
                    case "style:page-layout-properties":
                        this.ReadPageLayoutProps(props);
                        break;

                    default:
                        throw new ArgumentException("Invalid properties type");
                }
        }

        public void WriteStyle(XmlWriter xml)
        {
            if (!SetId())
                return;

            xml.WriteStartElement("style:page-layout");
            xml.WriteAttributeString("style:name", StyleId);

            this.WritePageLayoutProps(xml);

            xml.WriteEndElement();
        }
    }
}
