using System;
using System.IO;
using System.Diagnostics;

using ICSharpCode.SharpZipLib.Zip;


public interface IOpenDocument
{
    string Mimetype {get;}
    string Body {get;}
    string Style {get;}
}


public static class OpenDocumentExtension
{
    public static void Write(this IOpenDocument doc, string filePath)
    {
        var content = Resources.Get("content.xml", doc.Style, doc.Body);
        var manifest = Resources.Get("manifest.xml", doc.Mimetype);

        File.Delete(filePath);
        using (var zip = new ZipOutputStream(File.Create(filePath)))
        {
            zip.UseZip64 = UseZip64.Off;

            Add(zip, "mimetype", doc.Mimetype);
            Add(zip, "content.xml", content);
            Add(zip, "META-INF/manifest.xml", manifest);

            zip.Finish();
            zip.Close();
        }
    }

    private static void Add(ZipOutputStream zip, string name, string data)
    {
        var entry = new ZipEntry(name);

        if (name == "mimetype")
            entry.CompressionMethod = CompressionMethod.Stored;

        zip.PutNextEntry(entry);

        var writer = new StreamWriter(zip);
        writer.Write(data);

        writer.Flush();
        zip.CloseEntry();
    }
}