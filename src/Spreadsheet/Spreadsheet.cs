using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

using Genodf.Styles;

namespace Genodf
{
    public class Spreadsheet : IOpenDocument
    {
        public List<Sheet> Sheets { get; private set; }
        public List<IStyleable> Styles { get; private set; }

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

                foreach (var table in Sheets)   
                    WriteSheet(xml, table);

                xml.WriteEndElement();  // </office:spreadsheet>

                return contents.ToString();
            }
        }

        public string GlobalStyle
        {
            get
            {
                var builder = new StringBuilder();
                var writer = new StringWriter(builder);

                var xml = new XmlTextWriter(writer);
                xml.Formatting = Formatting.Indented;

                xml.WriteStartElement("office:styles");

                foreach (var style in Styles)
                {
                    style.WriteStyle(xml);
                }

                xml.WriteEndElement();

                return builder.ToString();
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

                // Generate empty styles
                (new Column()).WriteStyle(xml);
                (new Cell(0, 0)).WriteStyle(xml);

                foreach (var sheet in Sheets)
                {
                    foreach (var column in sheet.Columns)
                        column.WriteStyle(xml);

                    foreach (var row in sheet.Rows)
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
            Sheets = new List<Sheet>();
            Styles = new List<IStyleable>();
        }

        public Sheet NewSheet(string name)
        {
            var sheet = new Sheet(name);
            Sheets.Add(sheet);
            return sheet;
        }

        public void AddGlobalStyle(IStyleable style)
        {
            if (string.IsNullOrEmpty(style.Name))
                throw new ArgumentNullException("style.Name");
            Styles.Add(style);
        }

        public void AddGlobalStyle(string name, IStyleable style)
        {
            style.Name = name;
            AddGlobalStyle(style);
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

        private void WriteSheet(XmlWriter xml, Sheet sheet)
        {
            xml.WriteStartElement("table:table");
            xml.WriteAttributeString("table:name", sheet.Name);

            if (!sheet.Columns.Any())
                xml.WriteElementString("table:table-column", "");

            foreach (var column in sheet.Columns)
                column.Write(xml);

            if (!sheet.Rows.Any())
            {
                xml.WriteStartElement("table:table-row");
                xml.WriteElementString("table:table-cell", "");
                xml.WriteEndElement();
            }

            foreach (var row in sheet.Rows)
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

    public class Sheet
    {
        public string Name = "Sheet";

        public List<List<Cell>> Rows { get; private set; }
        public List<Column> Columns { get; private set; }

        public Sheet()
        {
            Rows = new List<List<Cell>>();
            Columns = new List<Column>();
        }

        public Sheet(string name) : this()
        {
            Name = name;
        }

        public Cell SetCell(string a1, string value)
        {
            int col, row;
            Spreadsheet.FromA1(a1, out col, out row);
            while (row >= Rows.Count)
                Rows.Add(new List<Cell>());

            while (col >= Rows[row].Count)
                Rows[row].Add(null);

            Rows[row][col] = new Cell(col, row, value);
            return Rows[row][col];
        }

        public Cell SetCell(int column, int row, string value)
        {
            return this.SetCell(Spreadsheet.ToA1(column, row), value);
        }

        public void SetColumn(int column)
        {
            for (int y = Columns.Count; column >= Columns.Count; y++)
                Columns.Add(new Column());
        }

        public Cell GetCell(string a1)
        {
            int col, row;
            Spreadsheet.FromA1(a1, out col, out row);

            if (row < Rows.Count)
                if (col < Rows[row].Count)
                    return Rows[row][col] ?? (Rows[row][col] = new Cell(col, row));

            this.SetCell(a1, string.Empty);
            return this.GetCell(a1);
        }

        public Cell GetCell(int column, int row)
        {
            return this.GetCell(Spreadsheet.ToA1(column, row));
        }

        public List<Cell> GetCells(string a1)
        {
            var topLeftA1 = a1.Split(':')[0];
            var bottomRightA1 = a1.Split(':')[1];

            int column, row;
            Spreadsheet.FromA1(topLeftA1, out column, out row);

            int tmp1, tmp2;
            Spreadsheet.FromA1(bottomRightA1, out tmp1, out tmp2);

            int width = tmp1 - column + 1;
            int height = tmp2 - row + 1;

            return GetCells(column, row, width, height);
        }

        public List<Cell> GetCells(int column, int row, int width, int height)
        {
            var cells = new List<Cell>();

            for (var x = column; x < column + width; x++)
                for (var y = row; y < row + height; y++)
                    cells.Add(GetCell(x, y));

            return cells;
        }

        public Column GetColumn(int column)
        {
            if (column < Columns.Count)
                return Columns[column];

            this.SetColumn(column);
            return this.GetColumn(column);
        }

        public Column GetColumn(string a1)
        {
            int column, row;
            Spreadsheet.FromA1(a1 + "1", out column, out row);
            return GetColumn(column);
        }
    }
}
