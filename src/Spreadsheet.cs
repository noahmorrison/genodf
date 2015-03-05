using System.Collections.Generic;
using System.Text;
using System.Xml;

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
            var settings = new XmlWriterSettings();
            settings.ConformanceLevel = ConformanceLevel.Fragment;
            var xml = XmlWriter.Create(contents, settings);

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
            return "";
        }
    }

    public Spreadsheet()
    {
        rows = new List<List<Cell>>();
        var row = new List<Cell>();
        row.Add(new Cell("1"));
        row.Add(new Cell("2"));
        rows.Add(row);

        row = new List<Cell>();
        row.Add(new Cell("=SUM(A1:C1)"));
        rows.Add(row);
    }

    private void WriteTable(XmlWriter xml)
    {
        xml.WriteStartElement("table:table");
        xml.WriteAttributeString("table:name", Name);

        xml.WriteElementString("table:table-column", null);
        foreach (var row in rows)
        {
            xml.WriteStartElement("table:table-row");
            foreach (var cell in row)
            {
                cell.Write(xml);
            }
            xml.WriteEndElement();  // </table:table-row>
        }
        xml.WriteEndElement();  // </table:table>
    }
}

class Cell
{
    private string value;

    public Cell()
    {
        value = "";
    }

    public Cell(string value) : this()
    {
        this.value = value;
    }

    public void Write(XmlWriter xml)
    {
        xml.WriteStartElement("table:table-cell");
        if (value[0] == '=')
        {
            xml.WriteAttributeString("office:value-type", "float");
            xml.WriteAttributeString("office:value", "0.00");
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
                xml.WriteAttributeString("office:value", value);
            }
        }
        xml.WriteEndElement();  // </table:table-cell>
    }
}
