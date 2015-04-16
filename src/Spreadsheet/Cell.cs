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
        public string ValueType;

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

            var type = "";
            if (string.IsNullOrEmpty(ValueType))
            {
                double tmp;
                if (value[0] == '=')
                    type = "function";

                else if (double.TryParse(value, out tmp))
                    type = "float";

                else
                    type = "string";
            }
            else
                type = ValueType;

            xml.WriteAttributeString("office:value-type", type);
            switch (type)
            {
                case "function":
                    xml.WriteAttributeString("table:formula", "of:" + value);
                    break;

                case "float":
                    xml.WriteAttributeString("office:value", value);
                    break;

                case "string":
                    xml.WriteElementString("text:p", value);
                    break;

                default:
                    break;
            }

            xml.WriteEndElement();  // </table:table-cell>
        }
    }
}