using System.Xml;

public interface IFormat
{
    string Code { get; }
    string Name { get; }
    void WriteFormat(XmlWriter xml);
}
