using System.Xml;

namespace Genodf
{
    public class Cell : ITableCellProperties,
                    IParagraphProperties,
                    ITextProperties
    {
        public string value;
        public int Column { get; private set; }
        public int Row { get; private set; }
        public int SpannedRows;
        public int SpannedColumns;

        public string Bg { get; set; }
        public string Fg { get; set; }
        public bool Bold { get; set; }
        public string TextAlign { get; set; }
        public bool Border { get; set; }
        public bool BorderTop { get; set; }
        public bool BorderBottom { get; set; }
        public bool BorderLeft { get; set; }
        public bool BorderRight { get; set; }

        public Cell(int column, int row)
        {
            Column = column;
            Row = row;
        }

        public Cell(string a1)
        {
            int col, row;
            Spreadsheet.FromA1(a1, out col, out row);
            Column = col;
            Row = row;
        }

        public Cell(int column, int row, string value)
            : this(column, row)
        {
            this.value = value;
        }

        public Cell(string a1, string value)
            : this(a1)
        {
            this.value = value;
        }

        public string ToA1()
        {
            return Spreadsheet.ToA1(Column, Row);
        }

        public bool IsStyled()
        {
            return this.TableCellIsStyled() ||
                   this.ParagraphIsStyled() ||
                   this.TextIsStyled();
        }

        public void WriteStyle(XmlWriter xml)
        {
            if (!this.IsStyled())
                return;

            xml.WriteStartElement("style:style");
            xml.WriteAttributeString("style:family", "table-cell");
            xml.WriteAttributeString("style:name", "CS-" + this.ToA1());

            this.WriteTableCellProps(xml);
            this.WriteParagraphProps(xml);
            this.WriteTextProps(xml);

            xml.WriteEndElement();
        }

        public void Write(XmlWriter xml)
        {
            xml.WriteStartElement("table:table-cell");

            if (this.IsStyled())
                xml.WriteAttributeString("table:style-name", "CS-" + this.ToA1());

            if (SpannedRows > 1)
                xml.WriteAttributeString("table:number-rows-spanned", SpannedRows.ToString());
            if (SpannedColumns > 1)
                xml.WriteAttributeString("table:number-columns-spanned", SpannedColumns.ToString());

            if (string.IsNullOrEmpty(value))
            {
                xml.WriteEndElement();
                return;
            }

            if (value[0] == '=')
            {
                xml.WriteAttributeString("table:formula", "of:" + value);
            }
            else
            {
                double tmp;
                if (double.TryParse(value, out tmp))
                {
                    xml.WriteAttributeString("office:value-type", "float");
                    xml.WriteAttributeString("office:value", tmp.ToString());
                }
                else
                {
                    xml.WriteAttributeString("office:value-type", "string");
                    xml.WriteElementString("text:p", value);
                }
            }
            xml.WriteEndElement();  // </table:table-cell>
        }
    }
}