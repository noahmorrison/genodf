using System;
using System.IO;
using System.Diagnostics;


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
        var dir = Path.GetDirectoryName(filePath);
        var tmpRoot = Path.Combine(dir, "tmp-odf-root");
        var metadir = Path.Combine(tmpRoot, "META-INF");

        var content = Resources.Get("content.xml", doc.Style, doc.Body);
        var manifest = Resources.Get("manifest.xml", doc.Mimetype);

        if (Directory.Exists(tmpRoot))
            Directory.Delete(tmpRoot, true);

        Directory.CreateDirectory(tmpRoot);
        Directory.CreateDirectory(metadir);

        Add(tmpRoot, "mimetype", doc.Mimetype);
        Add(tmpRoot, "content.xml", content);
        Add(metadir, "manifest.xml", manifest);

        File.Delete(filePath);
        doc.Zip(tmpRoot, filePath);

        // Remove the temporary root
        Directory.Delete(tmpRoot, true);
    }

#if __MonoCS__
    private static void Zip(this IOpenDocument doc, string tmpRoot, string filePath)
    {
        var command = new ProcessStartInfo
        {
            FileName = "zip",
            WorkingDirectory = tmpRoot,
        };

        // Zip the mimetype
        command.Arguments = "-q -0 -X " + filePath + " mimetype";
        Process.Start(command).WaitForExit();

        // And everything else
        command.Arguments = "-q -r " + filePath + " . -x mimetype";
        Process.Start(command).WaitForExit();
    }
#else
    private static void Zip(this IOpenDocument doc, string filePath)
    {
        throw new NotImplementedException("You cannot yet zip the ODF on Windows");
    }
#endif

    private static void Add(string root, string name, string data)
    {
        var path = Path.Combine(root, name);
        var fs = new FileStream(path, FileMode.OpenOrCreate);
        var writer = new StreamWriter(fs);
        writer.Write(data);
        writer.Flush();

        fs.Close();
    }
}
