using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Genodf.Styles
{
    public class MasterPageStyle : Base,
        IStyleable
    {
        public PageLayoutStyle PageLayout;

        public MasterPageStyle(string name)
        {
            StyleName = name;
        }

        public void WriteStyle(XmlWriter xml)
        {
            if (!SetId())
                return;

            xml.WriteStartElement("style:master-page");

            xml.WriteAttributeString("style:name", this.StyleId);
            xml.WriteAttributeString("style:page-layout-name", PageLayout?.StyleName ?? "default");

            xml.WriteEndElement();
        }
    }
}
