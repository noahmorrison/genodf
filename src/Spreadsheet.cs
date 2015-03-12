using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Linq;

public class Spreadsheet : IOpenDocument
{
    private string Name = "test";
    private List<List<Cell>> rows;

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
    }

    public void SetCell(string a1, string value)
    {
        var cell = new Cell(a1, value);
        int x = cell.Column;
        int y = cell.Row;
        while (y >= rows.Count)
        {
            rows.Add(new List<Cell>());
        }

        var row = rows[y];
        while (x >= row.Count)
        {
            row.Add(null);
        }

        row[x] = cell;
    }

    public Cell GetCell(string a1)
    {
        var cell = new Cell(a1);
        var x = cell.Column;
        var y = cell.Row;

        if (y < rows.Count)
            if (x < rows[y].Count)
                return rows[y][x];
        return null;
    }

    private void WriteTable(XmlWriter xml)
    {
        xml.WriteStartElement("table:table");
        xml.WriteAttributeString("table:name", Name);

        xml.WriteElementString("table:table-column", null);
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
    private string value;
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
        string numbers = new string(a1.SkipWhile(
                                l => char.IsLetter(l)).ToArray());
        string letters = new string(a1.TakeWhile(
                                l => char.IsLetter(l)).ToArray());

        Row = int.Parse(numbers) - 1;
        Column = 0;
        if (letters != null && letters.Length > 0)
        {
            Column = letters[0] - 'A';
            for (int i = 1; i < letters.Length; i++)
            {
                Column *= 26;
                Column += letters[i] - 'A';
            }
        }
    }

    public Cell(int column, int row, string value) : this(column, row)
    {
        this.value = value;
    }

    public Cell(string a1, string value) : this(a1)
    {
        this.value = value;
    }

    public string ToA1()
    {
        string a1 = (Row + 1).ToString();
        int tmp = Column + 1;
        while (--tmp >= 0)
        {
            a1 = (char)('A' + tmp % 26) + a1;
            tmp /= 26;
        }
        return a1;
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