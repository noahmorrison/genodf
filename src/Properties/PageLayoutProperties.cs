using System;
using System.Xml;

namespace Genodf
{
    public interface IPageLayoutProperties
    {
        string PageWidth { get; set; }
        string PageHeight { get; set; }
        int? FitInPages { get; set; }
    }

    internal static class PageLayoutProperties
    {
        internal static void WritePageLayoutProps(this IPageLayoutProperties props, XmlWriter xml)
        {
            if (!PageIsStyled(props))
                return;

            xml.WriteStartElement("style:page-layout-properties");

            if (!string.IsNullOrEmpty(props.PageWidth))
                xml.WriteAttributeString("fo:page-width", props.PageWidth);
            if (!string.IsNullOrEmpty(props.PageHeight))
                xml.WriteAttributeString("fo:page-height", props.PageHeight);
            if (props.FitInPages.HasValue)
                xml.WriteAttributeString("style:scale-to-pages", props.FitInPages.ToString());

            xml.WriteEndElement();
        }

        internal static void ReadPageLayoutProps(this IPageLayoutProperties self, XmlNode node)
        {
            if (node.Name != "style:page-layout-properties")
                throw new ArgumentException("Xml node is not a table property node", "node");

            node.Attributes.IfHas("fo:page-width", value => self.PageWidth = value);
            node.Attributes.IfHas("fo:page-height", value => self.PageHeight = value);
            node.Attributes.IfHas("style:scale-to-pages", value => self.FitInPages = int.Parse(value));
        }

        internal static bool PageIsStyled(this IPageLayoutProperties props)
        {
            return !string.IsNullOrEmpty(props.PageWidth)
                || !string.IsNullOrEmpty(props.PageHeight)
                || props.FitInPages.HasValue;
        }
    }
}
