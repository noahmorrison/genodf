using System.Collections.Generic;
using System.Text;
using System.Linq;
using System;

namespace Genodf
{
    public class XmlWriter
    {
        private StringBuilder builder = new StringBuilder();
        private List<ElementWriter> elements = new List<ElementWriter>();
        private Stack<ElementWriter> elementStack = new Stack<ElementWriter>();

        public override string ToString()
        {
            var sb = new StringBuilder();
            foreach (var element in elements)
            {
                sb.Append(element.ToString());
            }

            return sb.ToString();
        }

        public void WriteStartElement(string name)
        {
            var element = new ElementWriter(name);

            if (elementStack.Any()) {
                elementStack.Peek().AddChild(element);
            } else {
                elements.Add(element);
            }

            elementStack.Push(element);
        }
        public void WriteAttributeString(string name, string value)
        {
            if (elementStack.Any()) {
                elementStack.Peek().AddAttribute(name, value);
            } else {
                elements.Last().AddAttribute(name, value);
            }
        }

        public void WriteElementString(string name, string value)
        {
            var element = new ElementWriter(name, value);

            if (elementStack.Any()) {
                elementStack.Peek().AddChild(element);
            } else {
                elements.Add(element);
            }
        }

        public void WriteEndElement()
        {
            if (elementStack.Any()) {
                elementStack.Pop();
            }
        }

        public void WriteValue(string value)
        {
            if (elementStack.Any()) {
                elementStack.Peek().AddValue(value);
            } else {
                elements.Last().AddValue(value);
            }
        }
    }

    public class ElementWriter
    {
        private string _name;
        private string _value;
        private List<(string, string)> _attributes = new List<(string, string)>();

        private List<ElementWriter> _children = new List<ElementWriter>();

        public ElementWriter(string name)
        {
            _name = name;
            _value = "";
        }

        public ElementWriter(string name, string value)
        {
            _name = name;
            _value = value;
        }

        public void AddAttribute(string name, string value)
        {
            _attributes.Add((name, value));
        }

        public void AddChild(ElementWriter child)
        {
            _children.Add(child);
        }

        public void AddValue(string value)
        {
            _value += value;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("<");
            sb.Append(_name);
            sb.Append(" ");
            foreach (var (name, value) in _attributes)
            {
                sb.Append(name);
                sb.Append("=\"");
                sb.Append(escape(value));
                sb.Append("\" ");
            }

            if (string.IsNullOrWhiteSpace(_value) && !_children.Any())
            {
                sb.Append("/>");
                return sb.ToString();
            }

            sb.Remove(sb.Length - 1, 1);
            sb.Append(">");

            foreach (var child in _children)
            {
                sb.Append(child);
            }

            sb.Append(escape(_value));

            sb.Append("</");
            sb.Append(_name);
            sb.Append(">");

            return sb.ToString();
        }

        private static string escape(string value)
        {
            return ((value ?? "")
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;")
            );
        }
    }
}
