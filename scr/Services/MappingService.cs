using System.IO;
using System.Xml;
using TwinCATModuleTransfer.Models;

namespace TwinCATModuleTransfer.Services
{
    public static class MappingService
    {
        public static string RewriteMappingXml(string mappingXml, MappingRewriteOptions options)
        {
            if (options == null ||
                string.IsNullOrWhiteSpace(options.OldPlcBasePath) ||
                string.IsNullOrWhiteSpace(options.NewPlcBasePath))
                return mappingXml;

            return mappingXml.Replace(options.OldPlcBasePath, options.NewPlcBasePath);
        }

        public static void Save(string path, string xml) { File.WriteAllText(path, xml); }
        public static string Load(string path) { return File.ReadAllText(path); }

        public static string Pretty(string xml)
        {
            var xd = new XmlDocument();
            xd.LoadXml(xml);
            using (var sw = new StringWriter())
            {
                using (var xw = new XmlTextWriter(sw))
                {
                    xw.Formatting = Formatting.Indented;
                    xd.Save(xw);
                }
                return sw.ToString();
            }
        }
    }
}