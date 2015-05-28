using System.Xml;
using Genodf.Styles;

namespace Genodf
{
    public class Cell : CellStyle
    {
        public string value;
        public int Column { get; private set; }
        public int Row { get; private set; }
        public int SpannedRows;
        public int SpannedColumns;
        public string ValueType;

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

        public void Write(XmlWriter xml)
        {
            xml.WriteStartElement("table:table-cell");

            xml.WriteAttributeString("table:style-name", StyleId);

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

                else if (Format != null && Format.Code.EndsWith("%"))
                    type = "percentage";

                else if (double.TryParse(value, out tmp))
                    type = "float";

                else
                    type = "string";
            }
            else
                type = ValueType;

            if (type != "function")
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

                case "percentage":
                    xml.WriteAttributeString("office:value", value);
                    break;

                default:
                    break;
            }

            xml.WriteEndElement();  // </table:table-cell>
        }
    }
}
