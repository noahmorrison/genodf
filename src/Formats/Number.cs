using System;
using System.Xml;
using System.Linq;
using System.Collections.Generic;

using Genodf.Styles;

namespace Genodf
{
    public class NumberFormat : IFormat,
        ITextProperties
    {
        #region Properties
        public bool Bold { get; set; }
        public string Fg { get; set; }
        public string Bg { get; set; }
        #endregion

        private static int _count = 0;
        private static Dictionary<string, string> _codes = new Dictionary<string, string>();

        private Dictionary<string, NumberFormat> conditions = new Dictionary<string, NumberFormat>();

        private int leadingZeros = 0;
        private int decimalPlaces = 0;
        private string prefix;
        private string suffix;

        public string Code { get; private set; }

        public string FormatId { get; private set; }

        public NumberFormat(string code)
        {
            Code = code;
        }

        public NumberFormat(XmlNode node, Func<string, XmlNode> search)
        {
            var intRead = false;
            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.Name == "style:text-properties")
                    this.ReadTextProps(child);

                if (child.Name == "number:text")
                    if (intRead)
                        suffix = child.InnerText;
                    else
                        prefix = child.InnerText;

                if (child.Name == "number:number")
                {
                    intRead = true;
                    if (child.Attributes["number:decimal-places"] != null)
                        decimalPlaces = int.Parse(child.Attributes["number:decimal-places"].Value);
                    if (child.Attributes["number:min-integer-digits"] != null)
                        leadingZeros = int.Parse(child.Attributes["number:min-integer-digits"].Value);
                }

                if (child.Name == "style:map")
                {
                    var cond = child.Attributes["style:condition"].Value;
                    var style = child.Attributes["style:apply-style-name"].Value;
                    var term = "number:number-style[@style:name = \"" + style + "\"]";
                    var styleNode = search(term);
                    conditions.Add(cond, new NumberFormat(styleNode, search));
                }
            }

            Code = prefix + new string('0', leadingZeros);
            if (decimalPlaces > 0)
                Code += "." + new string('0', decimalPlaces);
            Code += suffix;
        }

        internal static void Reset()
        {
            _count = 0;
            _codes = new Dictionary<string, string>();
        }

        public override string ToString()
        {
            var str = "[" + Code;
            foreach (var condition in conditions)
                str += " (" + condition.Key + condition.Value.ToString() + ")";
            return str + "]";
        }

        public void WriteFormat(XmlWriter xml)
        {
            if (!SetId())
                return;

            foreach (var condition in conditions)
                condition.Value.WriteFormat(xml);

            if (Code.EndsWith("%"))
                xml.WriteStartElement("number:percentage-style");
            else
                xml.WriteStartElement("number:number-style");

            xml.WriteAttributeString("style:name", FormatId);

            this.WriteTextProps(xml);

            if (!string.IsNullOrEmpty(prefix))
                xml.WriteElementString("number:text", prefix);

            xml.WriteStartElement("number:number");
            xml.WriteAttributeString("number:decimal-places", decimalPlaces.ToString());
            xml.WriteAttributeString("number:min-integer-digits", leadingZeros.ToString());
            xml.WriteEndElement();

            if (!string.IsNullOrEmpty(suffix))
                xml.WriteElementString("number:text", suffix);

            foreach (var condition in conditions)
            {
                xml.WriteStartElement("style:map");
                xml.WriteAttributeString("style:condition", condition.Key);
                xml.WriteAttributeString("style:apply-style-name", condition.Value.FormatId);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        private bool SetId()
        {
            if (!ParseFormat(Code))
                return false;

            var key = this.ToString();

            if (_codes.ContainsKey(key))
            {
                FormatId = _codes[key];
                return false;
            }
            else
            {
                var styleType = this.GetType().Name;

                FormatId = "NumFormat-" + _count;
                _codes[key] = FormatId;

                _count++;
                return true;
            }
        }

        public NumberFormat AddCondition(string code, string condition)
        {
            var numFormat = new NumberFormat(code);
            conditions.Add(condition, numFormat);
            return numFormat;
        }

        private bool ParseFormat(string code)
        {
            System.Console.WriteLine("Code: " + code);

            leadingZeros = 0;
            decimalPlaces = 0;

            var leadingCode = string.Empty;
            var decimalCode = string.Empty;

            if (!code.Contains("."))
                leadingCode = code;
            else if (code.Split('.').Length > 2)
                return false;
            else
            {
                leadingCode = code.Split('.')[0];
                decimalCode = code.Split('.')[1];
            }

            if (leadingCode.LastIndexOf('#') > leadingCode.IndexOf('0')
                && leadingCode.IndexOf('0') != -1)
                return false;
            if (decimalCode.LastIndexOf('0') > decimalCode.IndexOf('#')
                && decimalCode.IndexOf('#') != -1)
                return false;

            foreach (char c in leadingCode)
                if (c == '0')
                    leadingZeros++;

            foreach (char c in decimalCode)
                if (c == '0')
                    decimalPlaces++;

            prefix = new string(leadingCode.TakeWhile(
                    c => c != '0' && c != '#' && c != ',').ToArray());

            var ending = decimalCode;
            if (string.IsNullOrEmpty(ending))
                ending = leadingCode;

            suffix = new string(ending.Reverse().TakeWhile(
                    c => c != '0' && c != '#' && c != ',')
                                .Reverse().ToArray());
            return true;
        }
    }
}
