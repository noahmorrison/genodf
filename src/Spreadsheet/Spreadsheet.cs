using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;


namespace Genodf
{
    public class Spreadsheet : IOpenDocument
    {
        public string Name = "test";
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
                    return rows[row][col] ?? (rows[row][col] = new Cell(col, row));

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
            FromA1(topLeftA1, out column, out row);

            int tmp1, tmp2;
            FromA1(bottomRightA1, out tmp1, out tmp2);

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
            if (column < columns.Count)
                return columns[column];

            this.SetColumn(column);
            return this.GetColumn(column);
        }

        public Column GetColumn(string a1)
        {
            int column, row;
            FromA1(a1 + "1", out column, out row);
            return GetColumn(column);
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
}
