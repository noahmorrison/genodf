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
        public MasterPageStyle MasterPageStyle { get; private set; }

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

                xml.WriteStartElement("style:default-style");
                xml.WriteAttributeString("style:family", "table-cell");
                xml.WriteStartElement("style:text-properties");
                xml.WriteAttributeString("style:font-name", "Arial");
                xml.WriteEndElement();
                xml.WriteEndElement();

                foreach (var style in Styles)
                    style.WriteStyle(xml);

                xml.WriteEndElement();

                xml.WriteStartElement("office:master-styles");
                MasterPageStyle.WriteStyle(xml);
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
                    sheet.WriteStyle(xml);
                    
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
            MasterPageStyle = new MasterPageStyle("Default");
        }

        public void Load(DocumentFiles ods)
        {
            if (ods.Mimetype != Mimetype)
                throw new InvalidDataException("Mimetype does not match");

            #region Content
            var ns = new XmlNamespaceManager(new NameTable());
            foreach (XmlAttribute attr in ods.Content.ChildNodes[1].Attributes)
                if (attr.Name.StartsWith("xmlns:"))
                    ns.AddNamespace(attr.Name.Replace("xmlns:", string.Empty), attr.Value);

            var styleRoot = ods.Content.DocumentElement.SelectSingleNode("//office:automatic-styles", ns);
            var root = ods.Content.DocumentElement.SelectSingleNode("//office:spreadsheet", ns);
            foreach (XmlNode table in root.SelectNodes("table:table", ns))
            {
                var sheet = NewSheet(table.Attributes["table:name"].Value);
                #region Columns
                foreach (XmlNode node in table.SelectNodes("table:table-column", ns))
                {
                    var col = new Column();
                    node.Attributes.IfHas("table:style-name", styleName =>
                    {
                        var xpath = string.Format("style:style[@style:name = \"{0}\"]", styleName);
                        var styleNode = styleRoot.SelectSingleNode(xpath, ns);
                        foreach (XmlNode props in styleNode.ChildNodes) switch (props.Name)
                        {
                            case "style:table-column-properties":
                                col.ReadTableColumnProps(props);
                                break;

                            default:
                                throw new InvalidDataException("Invalid properties type");
                        }
                    });

                    sheet.Columns.Add(col);
                }
                #endregion
                #region Rows
                var y = 0;
                foreach (XmlNode rowNode in table.SelectNodes("table:table-row", ns))
                {
                    var x = 0;
                    var row = new List<Cell>();
                    foreach (XmlNode node in rowNode.SelectNodes("table:table-cell", ns))
                    {
                        var cell = new Cell(x, y);

                        var valueType = node.Attributes["office:value-type"];
                        switch (valueType != null ? valueType.Value : null)
                        {
                            case "string":
                                if (node.ChildNodes.Count > 0)
                                    cell.Value = node.ChildNodes[0].InnerText;
                                break;

                            case "float":
                            case "percentage":
                                node.Attributes.IfHas("office:value", v => cell.Value = v);
                                break;

                            default:
                                node.Attributes.IfHas("table:formula", value =>
                                    cell.Value = value.Substring(3));
                                break;
                        }
                        cell.ValueType = valueType != null ? valueType.Value : null;

                        node.Attributes.IfHas("table:style-name", styleName =>
                        {
                            #region Cell style
                            var xpath = string.Format("style:style[@style:name = \"{0}\"]", styleName);
                            var styleNode = styleRoot.SelectSingleNode(xpath, ns);
                            if (styleNode != null)
                                foreach (XmlNode props in styleNode.ChildNodes) switch (props.Name)
                                {
                                    case "style:paragraph-properties":
                                        cell.ReadParagraphProps(props);
                                        break;
                                    case "style:table-cell-properties":
                                        cell.ReadTableCellProps(props);
                                        break;
                                    case "style:text-properties":
                                        cell.ReadTextProps(props);
                                        break;
                                    case "style:map":
                                        var cond = props.Attributes["style:condition"].Value;
                                        var style = props.Attributes["style:apply-style-name"].Value;

                                        cell.AddConditional(cond, style);
                                        break;

                                    default:
                                        throw new InvalidDataException("Invalid properties type");
                                }

                            node.Attributes.IfHas("table:number-rows-spanned", value =>
                                cell.SpannedRows = int.Parse(value));

                            node.Attributes.IfHas("table:number-columns-spanned", value =>
                                cell.SpannedColumns = int.Parse(value));

                            #region Cell format
                            if (styleNode != null)
                                if (styleNode.Attributes["style:data-style-name"] != null)
                                {
                                    var format = styleNode.Attributes["style:data-style-name"].Value;
                                    var autoStyles = ods.Content.DocumentElement.SelectSingleNode("//office:automatic-styles", ns);
                                    var globalStyles = ods.Style.DocumentElement.SelectSingleNode("//office:styles", ns);

                                    var xp = string.Format("number:number-style[@style:name = \"{0}\"] | " +
                                                           "number:percentage-style[@style:name = \"{0}\"]",
                                                           format);

                                    var formatNode = autoStyles.SelectSingleNode(xp, ns);
                                    if (formatNode == null)
                                        formatNode = globalStyles.SelectSingleNode(xp, ns);
                                    cell.Format = new NumberFormat(formatNode, search => { return autoStyles.SelectSingleNode(search, ns); });
                                }
                            #endregion
                            #endregion
                        });

                        node.Attributes.IfHas("table:number-columns-repeated", value =>
                        {
                            var count = int.Parse(value) - 1;
                            for (var i = 0; i < count; i++)
                                row.Add(cell);
                        });

                        row.Add(cell);
                        x++;
                    }

                    sheet.Rows.Add(row);
                    y++;
                }
                #endregion
            }
            #endregion

            #region Global Styles
            var globalStyle = ods.Style.DocumentElement.SelectSingleNode("//office:styles", ns);
            foreach (XmlNode node in globalStyle.SelectNodes("style:style", ns))
                AddGlobalStyle(new CellStyle(node));
            #endregion
        }

        public Sheet NewSheet(string name)
        {
            var sheet = new Sheet(name);
            Sheets.Add(sheet);

            return sheet;
        }

        public void AddGlobalStyle(IStyleable style)
        {
            if (string.IsNullOrEmpty(style.StyleName))
                throw new ArgumentNullException("style.Name");
            Styles.Add(style);
        }

        public void AddGlobalStyle(string name, IStyleable style)
        {
            style.StyleName = name;
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

        public static void FromA1Box(string a1, out int column, out int row, out int width, out int height)
        {
            var topLeftA1 = a1.Split(':')[0];
            var bottomRightA1 = a1.Split(':')[1];

            Spreadsheet.FromA1(topLeftA1, out column, out row);

            int tmp1, tmp2;
            Spreadsheet.FromA1(bottomRightA1, out tmp1, out tmp2);

            width = tmp1 - column + 1;
            height = tmp2 - row + 1;
        }

        private void WriteSheet(XmlWriter xml, Sheet sheet)
        {
            xml.WriteStartElement("table:table");
            xml.WriteAttributeString("table:name", sheet.Name);
            xml.WriteAttributeString("table:style-name", sheet.StyleId);

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

    public class Sheet : TableStyle
    {
        public string Name = "Sheet";

        public List<List<Cell>> Rows { get; private set; }
        public List<Column> Columns { get; private set; }

        public Sheet() : base()
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
            int column, row;
            Spreadsheet.FromA1(a1, out column, out row);
            return SetCell(column, row, value);
        }

        public Cell SetCell(int column, int row, string value)
        {
            while (row >= Rows.Count)
                Rows.Add(new List<Cell>());

            while (column >= Rows[row].Count)
                Rows[row].Add(null);

            if (Rows[row][column] != null)
                Rows[row][column].Value = value;
            else
                Rows[row][column] = new Cell(column, row, value);

            return Rows[row][column];
        }

        public void SetColumn(int column)
        {
            for (int y = Columns.Count; column >= Columns.Count; y++)
                Columns.Add(new Column());
        }

        public Cell GetCell(string a1)
        {
            int column, row;
            Spreadsheet.FromA1(a1, out column, out row);

            return GetCell(column, row);
        }

        public Cell GetCell(int column, int row)
        {
            if (row < Rows.Count)
                if (column < Rows[row].Count)
                    return Rows[row][column] ?? (Rows[row][column] = new Cell(column, row));

            SetCell(column, row, string.Empty);
            return GetCell(column, row);
        }

        public Cell GetCell(string a1, Action<Cell> action)
        {
            var cell = GetCell(a1);
            action(cell);
            return cell;
        }

        public Cell GetCell(int column, int row, Action<Cell> action)
        {
            var cell = GetCell(column, row);
            action(cell);
            return cell;
        }

        public List<Cell> GetCells(string a1)
        {
            int column, row, width, height;
            Spreadsheet.FromA1Box(a1, out column, out row, out width, out height);

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

        public void BorderAround(string a1)
        {
            int column, row, width, height;
            Spreadsheet.FromA1Box(a1, out column, out row, out width, out height);
            BorderAround(column, row, width, height);
        }

        public void BorderAround(int column, int row, int width, int height)
        {
            for (var x = column; x < column + width; x++)
            {
                GetCell(x, row).BorderTop = true;
                GetCell(x, row + height - 1).BorderBottom = true;
                if (GetCell(x, row).SpannedRows == height)
                    GetCell(x, row).BorderBottom = true;
            }

            for (var y = row; y < row + height; y++)
            {
                GetCell(column, y).BorderLeft = true;
                GetCell(column + width - 1, y).BorderRight = true;
                if (GetCell(column, y).SpannedColumns == width)
                    GetCell(column, y).BorderRight = true;
            }
        }
    }
}
