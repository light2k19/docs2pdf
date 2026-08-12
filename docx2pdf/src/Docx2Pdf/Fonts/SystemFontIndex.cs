using System;
using System.Collections.Generic;
using System.IO;

namespace Docx2Pdf.Fonts
{
    internal sealed class SystemFontEntry
    {
        public string Path;
        public int Index;
        public string Family;
        public string Subfamily;
        public bool Bold;
        public bool Italic;
        public ushort Weight;
    }

    /// <summary>Indexes the font files installed on the machine so document fonts can be embedded.</summary>
    internal static class SystemFontIndex
    {
        private static readonly object Gate = new object();
        private static Dictionary<string, List<SystemFontEntry>> _byFamily;
        private static readonly HashSet<string> ScannedDirectories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public static IEnumerable<string> DefaultDirectories()
        {
            var dirs = new List<string>();
            Action<string> add = d =>
            {
                if (!string.IsNullOrEmpty(d) && !dirs.Contains(d))
                    dirs.Add(d);
            };

            try { add(Environment.GetFolderPath(Environment.SpecialFolder.Fonts)); }
            catch { /* not available on this platform */ }

            var windir = Environment.GetEnvironmentVariable("WINDIR");
            if (!string.IsNullOrEmpty(windir))
                add(Path.Combine(windir, "Fonts"));

            var localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
            if (!string.IsNullOrEmpty(localAppData))
            {
                add(Path.Combine(localAppData, @"Microsoft\Windows\Fonts"));
                // Office's on-demand cloud fonts (Word lays documents out with these —
                // e.g. Lato): per-family subfolders holding numbered .ttf files.
                add(Path.Combine(localAppData, @"Microsoft\FontCache\4\CloudFonts"));
            }

            var home = Environment.GetEnvironmentVariable("HOME");
            add("/usr/share/fonts");
            add("/usr/local/share/fonts");
            add("/Library/Fonts");
            add("/System/Library/Fonts");
            add("/System/Library/Fonts/Supplemental");
            if (!string.IsNullOrEmpty(home))
            {
                add(Path.Combine(home, ".fonts"));
                add(Path.Combine(home, ".local/share/fonts"));
                add(Path.Combine(home, "Library/Fonts"));
            }
            return dirs;
        }

        public static void Build(IEnumerable<string> directories)
        {
            lock (Gate)
            {
                if (_byFamily == null)
                    _byFamily = new Dictionary<string, List<SystemFontEntry>>(StringComparer.OrdinalIgnoreCase);

                foreach (var dir in directories)
                {
                    if (string.IsNullOrEmpty(dir) || ScannedDirectories.Contains(dir))
                        continue;
                    ScannedDirectories.Add(dir);
                    ScanDirectory(dir);
                }
            }
        }

        private static void ScanDirectory(string dir)
        {
            string[] files;
            try
            {
                if (!Directory.Exists(dir))
                    return;
                files = Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories);
            }
            catch
            {
                return;
            }

            foreach (var file in files)
            {
                var ext = Path.GetExtension(file);
                if (string.IsNullOrEmpty(ext))
                    continue;
                ext = ext.ToLowerInvariant();
                if (ext != ".ttf" && ext != ".otf" && ext != ".ttc" && ext != ".otc")
                    continue;

                List<TrueTypeFile> probed;
                try
                {
                    probed = TrueTypeFile.Probe(file);
                }
                catch
                {
                    continue;
                }

                foreach (var font in probed)
                {
                    var families = new List<string>();
                    if (!string.IsNullOrEmpty(font.FamilyName)) families.Add(font.FamilyName);
                    if (!string.IsNullOrEmpty(font.TypographicFamily) && font.TypographicFamily != font.FamilyName)
                        families.Add(font.TypographicFamily);
                    if (families.Count == 0)
                        continue;

                    var entry = new SystemFontEntry
                    {
                        Path = font.FilePath,
                        Index = font.CollectionIndex,
                        Family = families[0],
                        Subfamily = font.SubfamilyName,
                        Bold = font.IsBold,
                        Italic = font.IsItalic,
                        Weight = font.WeightClass,
                    };

                    foreach (var family in families)
                    {
                        List<SystemFontEntry> list;
                        if (!_byFamily.TryGetValue(family, out list))
                        {
                            list = new List<SystemFontEntry>();
                            _byFamily[family] = list;
                        }
                        list.Add(entry);
                    }
                }
            }
        }

        public static SystemFontEntry Find(string family, bool bold, bool italic)
        {
            if (string.IsNullOrEmpty(family) || _byFamily == null)
                return null;

            List<SystemFontEntry> candidates;
            if (!_byFamily.TryGetValue(family.Trim(), out candidates) || candidates.Count == 0)
                return null;

            SystemFontEntry best = null;
            var bestScore = int.MaxValue;
            foreach (var entry in candidates)
            {
                var score = 0;
                if (entry.Bold != bold) score += 10;
                if (entry.Italic != italic) score += 10;
                var targetWeight = bold ? 700 : 400;
                score += Math.Abs(entry.Weight - targetWeight) / 100;
                // Prefer plain "Regular"/"Bold" faces over condensed/black variants of the same family.
                if (!string.IsNullOrEmpty(entry.Subfamily))
                {
                    var sub = entry.Subfamily.ToLowerInvariant();
                    if (sub.Contains("condensed") || sub.Contains("black") || sub.Contains("light")
                        || sub.Contains("semi") || sub.Contains("thin"))
                        score += 3;
                }
                if (score < bestScore)
                {
                    bestScore = score;
                    best = entry;
                }
            }
            return best;
        }

        public static bool HasFamily(string family)
        {
            return !string.IsNullOrEmpty(family) && _byFamily != null && _byFamily.ContainsKey(family.Trim());
        }

        public static int FamilyCount
        {
            get { return _byFamily == null ? 0 : _byFamily.Count; }
        }
    }
}
