using System.Xml;

public interface IFormat
{
    string Code { get; }
    string FormatId { get; }
    void WriteFormat(XmlWriter xml);
}

public interface IFormatable
{
    IFormat Format { get; set; }
}
