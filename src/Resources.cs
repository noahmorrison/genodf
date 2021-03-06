using System.IO;
using System.Reflection;

public static class Resources
{
    public static string[] Names()
    {
        return typeof(Resources)
            .GetTypeInfo()
            .Assembly
            .GetManifestResourceNames();
    }

    public static string Get(string name, params string[] args)
    {
        var assembly = typeof(Resources).GetTypeInfo().Assembly;

        using (Stream stream = assembly.GetManifestResourceStream(name))
        {
            if (stream == null)
                return string.Empty;

            using (StreamReader reader = new StreamReader(stream))
            {
                string res = reader.ReadToEnd();
                int num = 1;
                foreach(var arg in args)
                {
                    res = res.Replace("{{" + num.ToString() + "}}", arg);
                    num++;
                }
                return res;
            }
        }
    }
}
