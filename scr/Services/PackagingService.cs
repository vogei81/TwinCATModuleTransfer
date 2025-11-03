using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;            // <— wichtig für ZipArchive, ZipArchiveMode, ZipArchiveEntry
using System.Runtime.Serialization;      // DataContract, DataMember
using System.Runtime.Serialization.Json; // DataContractJsonSerializer
using System.Text;

namespace TwinCATModuleTransfer.Services
{
    /// <summary>
    /// Packt/Entpackt ein Modul-Paket (.tcmodpkg = ZIP) mit Manifest.
    /// Struktur:
    ///  /manifest.json
    ///  /items/<index>_<safeName>.xml
    ///  /child/<index>_<safeName>.childexport    (optional)
    ///  /_Mappings.xml                           (optional)
    /// </summary>
    public static class PackagingService
    {
        public const string PackageExtension = ".tcmodpkg";
        private const string ManifestFile = "manifest.json";
        private const string ItemsFolder = "items";
        private const string ChildFolder = "child";
        private const string MappingsFile = "_Mappings.xml";

        #region Manifest-DTOs

        [DataContract]
        public class PackageManifest
        {
            [DataMember(Order = 1)] public string FormatVersion = "1.0";
            [DataMember(Order = 2)] public string ExportedAtUtc;
            [DataMember(Order = 3)] public string TwinCATVersion;     // optional
            [DataMember(Order = 4)] public string Comment;            // optional
            [DataMember(Order = 5)] public List<ManifestItem> Items = new List<ManifestItem>();
            [DataMember(Order = 6)] public bool ContainsMappings;
        }

        [DataContract]
        public class ManifestItem
        {
            [DataMember(Order = 1)] public int Index;
            [DataMember(Order = 2)] public string Name;
            [DataMember(Order = 3)] public string TwinCatPath;   // ITcSmTreeItem.PathName (Quelle)
            [DataMember(Order = 4)] public string XmlFile;       // items/<file>
            [DataMember(Order = 5)] public string ChildFile;     // child/<file> (optional)
            [DataMember(Order = 6)] public string ItemGuid;      // optional (Stabilisierung/Verfolgung)
        }

        #endregion

        public static void CreatePackage(string packagePath, PackageManifest manifest, List<PackageFile> files)
        {
            if (File.Exists(packagePath)) File.Delete(packagePath);

            using (var fs = new FileStream(packagePath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create, false, Encoding.UTF8))
            {
                // Manifest schreiben
                var manifestEntry = zip.CreateEntry(ManifestFile, CompressionLevel.Optimal);
                using (var s = manifestEntry.Open())
                {
                    var json = Serialize(manifest);
                    var buf = Encoding.UTF8.GetBytes(json);
                    s.Write(buf, 0, buf.Length);
                }

                // Dateien schreiben
                foreach (var f in files)
                {
                    var entry = zip.CreateEntry(NormalizeZipPath(f.PackageRelativePath), CompressionLevel.Optimal);
                    using (var es = entry.Open())
                    using (var inFs = new FileStream(f.SourceFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        inFs.CopyTo(es);
                    }
                }
            }
        }

        public static UnpackedPackage OpenPackage(string packagePath, string tempWorkingDir)
        {
            if (!File.Exists(packagePath))
                throw new FileNotFoundException("Paketdatei nicht gefunden", packagePath);

            if (Directory.Exists(tempWorkingDir)) Directory.Delete(tempWorkingDir, true);
            Directory.CreateDirectory(tempWorkingDir);

            using (var fs = new FileStream(packagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read, false, Encoding.UTF8))
            {
                ZipArchiveEntry manifestEntry = zip.GetEntry(ManifestFile);
                if (manifestEntry == null)
                    throw new InvalidOperationException("manifest.json fehlt im Paket.");

                PackageManifest manifest;
                using (var ms = new MemoryStream())
                using (var s = manifestEntry.Open())
                {
                    s.CopyTo(ms);
                    var json = Encoding.UTF8.GetString(ms.ToArray());
                    manifest = Deserialize<PackageManifest>(json);
                }

                // Alle Dateien entpacken
                foreach (var e in zip.Entries)
                {
                    var target = Path.Combine(tempWorkingDir, e.FullName.Replace('/', Path.DirectorySeparatorChar));
                    var dir = Path.GetDirectoryName(target);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                    using (var src = e.Open())
                    using (var dst = new FileStream(target, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        src.CopyTo(dst);
                    }
                }

                var unpack = new UnpackedPackage
                {
                    TempFolder = tempWorkingDir,
                    Manifest = manifest,
                    MappingsPath = manifest.ContainsMappings ? Path.Combine(tempWorkingDir, MappingsFile) : null
                };

                return unpack;
            }
        }

        public class UnpackedPackage
        {
            public string TempFolder;
            public PackageManifest Manifest;
            public string MappingsPath;

            public string ResolveItemPath(string relative)
            {
                if (string.IsNullOrEmpty(relative)) return null;
                var local = relative.Replace('/', Path.DirectorySeparatorChar);
                return Path.Combine(TempFolder, local);
            }
        }

        public class PackageFile
        {
            public string SourceFilePath;      // Pfad in deinem Export-Temp
            public string PackageRelativePath; // z. B. items/001_Axis1.xml
        }

        private static string Serialize<T>(T obj)
        {
            var ser = new DataContractJsonSerializer(typeof(T), new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            });
            using (var ms = new MemoryStream())
            {
                ser.WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private static T Deserialize<T>(string json)
        {
            var ser = new DataContractJsonSerializer(typeof(T), new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            });
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)ser.ReadObject(ms);
            }
        }

        private static string NormalizeZipPath(string path)
        {
            // ZIP erwartet '/' als Separator
            return path.Replace('\\', '/');
        }
    }
}