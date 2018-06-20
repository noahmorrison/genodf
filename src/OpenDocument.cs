using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Genodf
{
    public interface IOpenDocument
    {
        string Mimetype { get; }
        string Body { get; }
        string Style { get; }
        string GlobalStyle { get; }

        string Settings { get; }

        void Load(DocumentFiles files);
    }

    public static class OpenDocument
    {
        public static DocType Read<DocType>(string filePath) where DocType : IOpenDocument, new()
        {
            using (var stream = File.OpenRead(filePath))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                string mimetype;
                DocumentFiles files;

                using (var entry = zip.GetEntry("mimetype").Open())
                using (var reader = new StreamReader(entry))
                {
                    mimetype = reader.ReadToEnd();
                }

                using (var content = zip.GetEntry("content.xml").Open())
                using (var style = zip.GetEntry("styles.xml").Open())
                {
                    files = new DocumentFiles(mimetype, content, style);
                }

                var odf = new DocType();
                odf.Load(files);
                return odf;
            }
        }
    }

    public static class OpenDocumentExtension
    {
        public static byte[] GetBytes(this IOpenDocument doc)
        {
            Genodf.Styles.Base.Reset();
            Genodf.NumberFormat.Reset();

            var style = Resources.Get("styles.xml", doc.GlobalStyle);
            var content = Resources.Get("content.xml", doc.Style, doc.Body);
            var manifest = Resources.Get("manifest.xml", doc.Mimetype);
            var settings = Resources.Get("settings.xml", doc.Settings);

            using (var stream = new MemoryStream())
            {
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
                {
                    Add(zip, "mimetype", doc.Mimetype, CompressionLevel.NoCompression);
                    Add(zip, "content.xml", content);
                    Add(zip, "styles.xml", style);
                    Add(zip, "META-INF/manifest.xml", manifest);
                    Add(zip, "settings.xml", settings);
                }

                stream.Seek(0, SeekOrigin.Begin);
                return stream.ToArray();
            }
        }

        public static void Write(this IOpenDocument doc, string filePath)
        {
            File.Delete(filePath);
            File.WriteAllBytes(filePath, doc.GetBytes());
        }

        private static void Add(ZipArchive zip, string name, string data, CompressionLevel level = CompressionLevel.Fastest)
        {
            using (var entry = zip.CreateEntry(name, level).Open())
            using (var writer = new StreamWriter(entry))
            {
                writer.Write(data);
            }
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
