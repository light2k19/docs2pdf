using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Docx2Pdf.Ooxml
{
    /// <summary>A single relationship declared by a package part.</summary>
    internal sealed class OpcRelationship
    {
        public string Id;
        public string Type;
        public string Target;
        public bool External;
    }

    /// <summary>
    /// Minimal read-only Open Packaging Conventions (OPC) reader built on <see cref="ZipArchive"/>.
    /// Part names are normalised to absolute, leading-slash-free paths (e.g. <c>word/document.xml</c>).
    /// </summary>
    internal sealed class OpcPackage : IDisposable
    {
        private readonly ZipArchive _zip;
        private readonly Dictionary<string, ZipArchiveEntry> _entries =
            new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Dictionary<string, OpcRelationship>> _relCache =
            new Dictionary<string, Dictionary<string, OpcRelationship>>(StringComparer.OrdinalIgnoreCase);

        private OpcPackage(ZipArchive zip)
        {
            _zip = zip;
            foreach (var e in _zip.Entries)
                _entries[Normalize(e.FullName)] = e;
        }

        public static OpcPackage Open(Stream stream, bool leaveOpen)
        {
            return new OpcPackage(new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen));
        }

        public IEnumerable<string> PartNames { get { return _entries.Keys; } }

        public bool HasPart(string partName)
        {
            return partName != null && _entries.ContainsKey(Normalize(partName));
        }

        public byte[] ReadPart(string partName)
        {
            ZipArchiveEntry entry;
            if (!_entries.TryGetValue(Normalize(partName), out entry))
                return null;

            using (var src = entry.Open())
            using (var ms = new MemoryStream())
            {
                src.CopyTo(ms);
                return ms.ToArray();
            }
        }

        public XDocument ReadXml(string partName)
        {
            var bytes = ReadPart(partName);
            if (bytes == null)
                return null;

            using (var ms = new MemoryStream(bytes))
            using (var reader = XmlReader.Create(ms, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreWhitespace = false,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
            }))
            {
                return XDocument.Load(reader, LoadOptions.None);
            }
        }

        /// <summary>Relationships declared by <paramref name="partName"/> (pass null/empty for the package root).</summary>
        public Dictionary<string, OpcRelationship> GetRelationships(string partName)
        {
            var key = partName ?? string.Empty;
            Dictionary<string, OpcRelationship> cached;
            if (_relCache.TryGetValue(key, out cached))
                return cached;

            var map = new Dictionary<string, OpcRelationship>(StringComparer.Ordinal);
            var relPart = RelationshipPartFor(key);
            var doc = ReadXml(relPart);
            if (doc != null && doc.Root != null)
            {
                foreach (var el in doc.Root.Elements(Ns.Rel + "Relationship"))
                {
                    var rel = new OpcRelationship
                    {
                        Id = (string)el.Attribute("Id"),
                        Type = (string)el.Attribute("Type"),
                        Target = (string)el.Attribute("Target"),
                        External = string.Equals((string)el.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase),
                    };
                    if (!string.IsNullOrEmpty(rel.Id))
                        map[rel.Id] = rel;
                }
            }

            _relCache[key] = map;
            return map;
        }

        public OpcRelationship GetRelationship(string partName, string id)
        {
            if (string.IsNullOrEmpty(id))
                return null;
            OpcRelationship rel;
            return GetRelationships(partName).TryGetValue(id, out rel) ? rel : null;
        }

        public OpcRelationship FindRelationshipByType(string partName, string type)
        {
            return GetRelationships(partName).Values.FirstOrDefault(
                r => string.Equals(r.Type, type, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Resolves a relationship target (which may be relative) against the owning part.</summary>
        public static string ResolveTarget(string ownerPartName, string target)
        {
            if (string.IsNullOrEmpty(target))
                return null;
            if (target.StartsWith("/", StringComparison.Ordinal))
                return Normalize(target);

            var dir = string.Empty;
            if (!string.IsNullOrEmpty(ownerPartName))
            {
                var idx = ownerPartName.Replace('\\', '/').LastIndexOf('/');
                if (idx >= 0)
                    dir = ownerPartName.Substring(0, idx + 1);
            }

            var combined = dir + target.Replace('\\', '/');

            // Collapse "." and ".." segments.
            var stack = new List<string>();
            foreach (var seg in combined.Split('/'))
            {
                if (seg.Length == 0 || seg == ".")
                    continue;
                if (seg == "..")
                {
                    if (stack.Count > 0)
                        stack.RemoveAt(stack.Count - 1);
                }
                else
                {
                    stack.Add(seg);
                }
            }
            return string.Join("/", stack);
        }

        private static string RelationshipPartFor(string partName)
        {
            if (string.IsNullOrEmpty(partName))
                return "_rels/.rels";

            var p = Normalize(partName);
            var idx = p.LastIndexOf('/');
            var dir = idx >= 0 ? p.Substring(0, idx + 1) : string.Empty;
            var file = idx >= 0 ? p.Substring(idx + 1) : p;
            return dir + "_rels/" + file + ".rels";
        }

        private static string Normalize(string name)
        {
            if (name == null)
                return string.Empty;
            var n = name.Replace('\\', '/');
            while (n.StartsWith("/", StringComparison.Ordinal))
                n = n.Substring(1);
            return n;
        }

        public void Dispose()
        {
            _zip.Dispose();
        }
    }
}
