using System;
using System.Xml;
using System.Linq;

namespace Genodf
{
    public class NumberFormat : IFormat
    {
        private static int _count = 0;
        private int id;

        private int leadingZeros = 0;
        private int decimalPlaces = 0;
        private string prefix;
        private string suffix;

        public string Code { get; private set; }
        public bool Valid { get; private set; }

        public string Name { get { return "N" +  id; } }

        public NumberFormat(string code)
        {
            _count++;
            id = _count;
            Code = code;
            Valid = ParseFormat(code);
        }

        public void WriteFormat(XmlWriter xml)
        {
            xml.WriteStartElement("number:number-style");
            xml.WriteAttributeString("style:name", Name);

            if (!string.IsNullOrEmpty(prefix))
                xml.WriteElementString("number:text", prefix);

            xml.WriteStartElement("number:number");
            xml.WriteAttributeString("number:decimal-places", decimalPlaces.ToString());
            xml.WriteAttributeString("number:min-integer-digits", leadingZeros.ToString());
            xml.WriteEndElement();

            if (!string.IsNullOrEmpty(suffix))
                xml.WriteElementString("number:text", suffix);

            xml.WriteEndElement();
        }

        private bool ParseFormat(string code)
        {
            System.Console.WriteLine("Code: " + code);

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
