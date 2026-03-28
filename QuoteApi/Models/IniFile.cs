using System;
using System.Collections.Generic;
using System.Text;

namespace QuoteApi.Models
{
    public class IniFile
    {
        public string Path { get; }

        public IniFile(string iniPath)
        {
            Path = iniPath;
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(iniPath) ?? "");
        }
        /// <summary>
        /// Read Data Value From the Ini File Encoding.UTF8
        /// </summary>
        /// <PARAM name="Section"></PARAM>
        /// <PARAM name="Key"></PARAM>
        /// <PARAM name="Path"></PARAM>
        /// <returns></returns>
        public string IniReadUTF8(string Section, string Key)
        {
            //StringBuilder temp = new StringBuilder(255);
            //int i = GetPrivateProfileString(Section, Key, "", temp,
            //                                255, this.path);
            //return temp.ToString();

            if (!File.Exists(this.Path)) return "";

            string[] lines = File.ReadAllLines(this.Path, Encoding.UTF8);
            string currentSection = "";

            foreach (string line in lines)
            {
                string trimmed = line.Trim();
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    currentSection = trimmed.Substring(1, trimmed.Length - 2);
                }
                else if (trimmed.Contains("=") && currentSection == Section)
                {
                    var parts = trimmed.Split('=', (char)2);
                    if (parts[0].Trim() == Key)
                    {
                        return parts[1].Trim();
                    }
                }
            }
            return "";
        }
        public void WriteValue(string section, string key, string value)
        {
            var lines = LoadLines();
            bool sectionFound = false;
            bool keyFound = false;
            string newLine = $"{key}={value.Trim()}";

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i].Trim();
                if (line.Equals($"[{section}]", StringComparison.OrdinalIgnoreCase))
                {
                    sectionFound = true;
                    // 找 key
                    int nextLine = i + 1;
                    while (nextLine < lines.Count)
                    {
                        var next = lines[nextLine].Trim();
                        if (next.StartsWith("[") || next.Trim().Length == 0)
                            break;

                        if (next.StartsWith($"{key}="))
                        {
                            lines[nextLine] = newLine;
                            keyFound = true;
                            break;
                        }
                        nextLine++;
                    }
                    if (!keyFound)
                    {
                        lines.Insert(nextLine, newLine);
                    }
                    break;
                }
            }

            if (!sectionFound)
            {
                lines.Add($"[{section}]");
                lines.Add(newLine);
            }

            File.WriteAllLines(Path, lines, Encoding.UTF8);
        }

        public string ReadValue(string section, string key)
        {
            var lines = LoadLines();
            string currentSection = "";

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    currentSection = trimmed[1..^1];
                }
                else if (trimmed.StartsWith($"{key}=") && currentSection == section)
                {
                    return trimmed[(key.Length + 1)..].Trim();
                }
            }
            return "";
        }

        public string[] GetSections()
        {
            var lines = LoadLines();
            var sections = new HashSet<string>();

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    sections.Add(trimmed[1..^1]);
                }
            }
            return sections.ToArray();
        }

        public string[] GetKeys(string section)
        {
            var lines = LoadLines();
            var keys = new List<string>();
            string currentSection = "";

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    currentSection = trimmed[1..^1];
                }
                else if (trimmed.Contains("=") && currentSection == section)
                {
                    var key = trimmed.Split('=', 2)[0].Trim();
                    keys.Add(key);
                }
            }
            return keys.ToArray();
        }

        public void DeleteKey(string section, string key)
        {
            var lines = LoadLines();
            string currentSection = "";
            bool inSection = false;

            for (int i = 0; i < lines.Count; i++)
            {
                var trimmed = lines[i].Trim();
                if (trimmed.Equals($"[{section}]", StringComparison.OrdinalIgnoreCase))
                {
                    currentSection = section;
                    inSection = true;
                }
                else if (trimmed.StartsWith("[") && inSection)
                {
                    inSection = false;
                }
                else if (inSection && trimmed.StartsWith($"{key}="))
                {
                    lines.RemoveAt(i);
                    break;
                }
            }

            File.WriteAllLines(Path, lines, Encoding.UTF8);
        }

        private List<string> LoadLines()
        {
            return File.Exists(Path)
                ? File.ReadAllLines(Path, Encoding.UTF8).ToList()
                : new List<string>();
        }
    }

}
