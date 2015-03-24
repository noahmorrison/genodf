using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Linq;


namespace Genodf
{
    public class Spreadsheet : IOpenDocument
    {
        private string Name = "test";
        private List<List<Cell>> rows;
        private List<Column> columns;

        public string Mimetype
        {
            get
            {
                return "application/vnd.oasis.opendocument.spreadsheet";
            }
        }
        public string Body
        {
            get
            {
                var contents = new StringBuilder();
                var writer = new StringWriter(contents);

                var xml = new XmlTextWriter(writer);
                xml.Formatting = Formatting.Indented;

                xml.WriteStartElement("office:spreadsheet");

                WriteTable(xml);

                xml.WriteEndElement();  // </office:spreadsheet>

                return contents.ToString();
            }
        }

        public string Style
        {
            get
            {
                var style = new StringBuilder();
                var writer = new StringWriter(style);

                var xml = new XmlTextWriter(writer);
                xml.Formatting = Formatting.Indented;

                foreach (var column in columns)
                    column.WriteStyle(xml);

                foreach (var row in rows)
                {
                    foreach (var cell in row)
                    {
                        if (cell == null)
                            continue;

                        cell.WriteStyle(xml);
                    }
                }

                return style.ToString();
            }
        }

        public Spreadsheet()
        {
            rows = new List<List<Cell>>();
            columns = new List<Column>();
        }

        public void SetCell(string a1, string value)
        {
            int col, row;
            Spreadsheet.FromA1(a1, out col, out row);
            while (row >= rows.Count)
                rows.Add(new List<Cell>());

            while (col >= rows[row].Count)
                rows[row].Add(null);

            rows[row][col] = new Cell(col, row, value);
        }

        public void SetCell(int column, int row, string value)
        {
            this.SetCell(Spreadsheet.ToA1(column, row), value);
        }

        public void SetColumn(int column)
        {
            for (int y = columns.Count; column >= columns.Count; y++)
                columns.Add(new Column(y));
        }

        public static string ToA1(int column, int row)
        {
            string a1 = (row + 1).ToString();
            int tmp = column + 1;
            while (--tmp >= 0)
            {
                a1 = (char)('A' + tmp % 26) + a1;
                tmp /= 26;
            }
            return a1;
        }

        public static void FromA1(string a1, out int col, out int row)
        {
            int.TryParse(new string(a1.ToCharArray()
                                      .SkipWhile(c => !char.IsDigit(c))
                                      .ToArray()), out row);
            row--;

            var a = new string(a1.ToCharArray()
                                 .TakeWhile(c => char.IsLetter(c))
                                 .ToArray());
            col = 0;
            foreach (char c in a.ToUpper())
                col = (26*col) + (c-'A') + 1;
            col--;
        }

        public Cell GetCell(string a1)
        {
            int col, row;
            Spreadsheet.FromA1(a1, out col, out row);

            if (row < rows.Count)
                if (col < rows[row].Count)
                    return rows[row][col];

            this.SetCell(a1, "");
            return this.GetCell(a1);
        }

        public Cell GetCell(int column, int row)
        {
            return this.GetCell(Spreadsheet.ToA1(column, row));
        }

        public Column GetColumn(int column)
        {
            if (column < columns.Count)
                return columns[column];

            this.SetColumn(column);
            return this.GetColumn(column);
        }

        private void WriteTable(XmlWriter xml)
        {
            xml.WriteStartElement("table:table");
            xml.WriteAttributeString("table:name", Name);

            foreach (var column in columns)
                column.Write(xml);

            foreach (var row in rows)
            {
                xml.WriteStartElement("table:table-row");

                bool hasChild = false;
                foreach (var cell in row)
                {
                    hasChild = true;
                    if (cell != null)
                        cell.Write(xml);
                    else
                        xml.WriteElementString("table:table-cell", null);
                }

                if (!hasChild)
                    xml.WriteElementString("table:table-cell", null);

                xml.WriteEndElement();  // </table:table-row>
            }
            xml.WriteEndElement();  // </table:table>
        }
    }

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

            if (value == null || value == "")
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

    public class Column : ITableColumnProperties
    {
        private int index;

        public double? Width { get; set; }

        public Column(int index)
        {
            this.index = index;
        }

        public bool IsStyled()
        {
            return this.TableColumnIsStyled();
        }

        public void WriteStyle(XmlWriter xml)
        {
            if (!this.IsStyled())
                return;

            xml.WriteStartElement("style:style");
            xml.WriteAttributeString("style:family", "table-column");
            xml.WriteAttributeString("style:name", "co" + this.index);

            this.WriteTableColumnProps(xml);

            xml.WriteEndElement();
        }

        public void Write(XmlWriter xml)
        {
            xml.WriteStartElement("table:table-column");

            if (this.IsStyled())
                xml.WriteAttributeString("table:style-name", "co" + this.index);

            xml.WriteEndElement();
        }
    }
}
