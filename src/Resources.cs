using System.IO;

public static class Resources
{
    public static string[] Names()
    {
        return System.Reflection.Assembly.GetExecutingAssembly()
                .GetManifestResourceNames();
    }

    public static string Get(string name, params string[] args)
    {
        var assembly = System.Reflection.Assembly.GetExecutingAssembly();

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
