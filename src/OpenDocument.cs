using System;
using System.IO;
using System.Text;
using System.Xml;

using ICSharpCode.SharpZipLib.Zip;


namespace Genodf
{
    public interface IOpenDocument
    {
        string Mimetype { get; }
        string Body { get; }
        string Style { get; }
        string GlobalStyle { get; }

        void Load(DocumentFiles files);
    }

    public static class OpenDocument
    {
        public static DocType Read<DocType>(string filePath) where DocType : IOpenDocument, new()
        {
            using (var fstream = File.OpenRead(filePath))
            using (var zip = new ZipFile(fstream))
            {
                string mimetype;

                using (var stream = zip.GetInputStream(zip.GetEntry("mimetype")))
                using (var reader = new StreamReader(stream))
                    mimetype = reader.ReadToEnd();

                var content = zip.GetInputStream(zip.GetEntry("content.xml"));
                var style = zip.GetInputStream(zip.GetEntry("styles.xml"));

                var files = new DocumentFiles(mimetype, content, style);

                var odf = new DocType();
                odf.Load(files);
                return odf;
            }
        }
    }

    public static class OpenDocumentExtension
    {
        public static void Write(this IOpenDocument doc, string filePath)
        {
            Genodf.Styles.Base.Reset();
            Genodf.NumberFormat.Reset();

            var style = Resources.Get("styles.xml", doc.GlobalStyle);
            var content = Resources.Get("content.xml", doc.Style, doc.Body);
            var manifest = Resources.Get("manifest.xml", doc.Mimetype);

            File.Delete(filePath);
            using (var zip = new ZipOutputStream(File.Create(filePath)))
            {
                zip.UseZip64 = UseZip64.Off;

                Add(zip, "mimetype", doc.Mimetype);
                Add(zip, "content.xml", content);
                Add(zip, "styles.xml", style);
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

    public class DocumentFiles
    {
        public string Mimetype { get; private set; }
        public XmlDocument Content { get; private set; }
        public XmlDocument Style { get; private set; }

        public DocumentFiles(string mimetype, Stream content, Stream style)
        {
            Mimetype = mimetype;
            Content = new XmlDocument();
            Content.Load(content);
            Style = new XmlDocument();
            Style.Load(style);
        }
    }

    internal static class CustomExtensionMethods
    {
        public static void IfHas(this XmlAttributeCollection self, string key, Action<string> action)
        {
            if (self[key] != null)
                action(self[key].Value);
        }
    }
}