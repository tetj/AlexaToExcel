using System.Text;
using HtmlAgilityPack;

namespace AlexaToExcel
{
    class HtmlToCsvConverter
    {
        /// <summary>
        /// Reads all &lt;table&gt; elements from the given HTML file and writes them
        /// into a single CSV file.  If the HTML is a Google Drive / Sheets wrapper
        /// that loads data via an iframe, the converter automatically follows the
        /// iframe src into the companion <c>_files</c> folder.
        /// </summary>
        public static int Convert(string htmlPath, string csvPath)
        {
            if (!File.Exists(htmlPath))
            {
                throw new FileNotFoundException($"HTML file not found: {htmlPath}");
            }

            // Resolve the actual data file (Google Drive saves data in a _files subfolder)
            string resolvedPath = ResolveDataFile(htmlPath);

            var doc = new HtmlDocument();
            doc.Load(resolvedPath, Encoding.UTF8);

            // Google Sheets uses <table class="waffle"> for the main data table
            var tables = doc.DocumentNode.SelectNodes("//table[contains(@class,'waffle')]")
                      ?? doc.DocumentNode.SelectNodes("//table");

            if (tables == null || tables.Count == 0)
            {
                throw new InvalidOperationException("No <table> elements found in the HTML file.");
            }

            int totalRows = 0;

            using var writer = new StreamWriter(csvPath, false, new UTF8Encoding(true));

            for (int t = 0; t < tables.Count; t++)
            {
                if (t > 0)
                {
                    writer.WriteLine(); // blank line between tables
                }

                var table = tables[t];

                // Collect all rows from <tbody> (skip <thead> — Google Sheets puts only
                // column-width shims there, not real headers)
                var bodyRows = table.SelectNodes(".//tbody/tr")
                            ?? table.SelectNodes(".//tr");

                if (bodyRows == null)
                {
                    continue;
                }

                foreach (var row in bodyRows)
                {
                    // Skip freezebar / separator rows injected by Google Sheets
                    if (row.InnerHtml.Contains("freezebar-cell", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // Only pick up <td> cells — skip <th> row-header cells (row numbers)
                    var cells = row.SelectNodes("./td");
                    if (cells == null || cells.Count == 0)
                    {
                        continue;
                    }

                    var values = new List<string>();
                    foreach (var cell in cells)
                    {
                        values.Add(EscapeCsvField(GetCellText(cell)));
                    }
                    writer.WriteLine(string.Join(",", values));
                    totalRows++;
                }
            }

            return totalRows;
        }

        /// <summary>
        /// If the HTML file is a Google Drive wrapper that uses an iframe to load
        /// the actual spreadsheet data, return the path to the inner HTML file.
        /// Otherwise return the original path unchanged.
        /// </summary>
        private static string ResolveDataFile(string htmlPath)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlPath, Encoding.UTF8);

            var iframe = doc.DocumentNode.SelectSingleNode("//iframe[@src]");
            if (iframe == null)
            {
                return htmlPath;
            }

            string src = iframe.GetAttributeValue("src", "");
            if (string.IsNullOrWhiteSpace(src))
            {
                return htmlPath;
            }

            // The src is relative (e.g. "./Name_files/sheet.html")
            string dir = Path.GetDirectoryName(htmlPath) ?? ".";
            string candidate = Path.GetFullPath(Path.Combine(dir, src));

            if (File.Exists(candidate))
            {
                Console.WriteLine($"  Resolved iframe → {candidate}");
                return candidate;
            }

            return htmlPath;
        }

        private static string GetCellText(HtmlNode cell)
        {
            // Prefer link text when the cell contains an <a> element
            var link = cell.SelectSingleNode(".//a");
            string raw;
            if (link != null)
            {
                raw = link.InnerText ?? "";
            }
            else
            {
                raw = cell.InnerText ?? "";
            }

            // Decode HTML entities and collapse whitespace
            var text = HtmlEntity.DeEntitize(raw);
            text = text.Replace("\r", " ").Replace("\n", " ");

            while (text.Contains("  "))
            {
                text = text.Replace("  ", " ");
            }

            return text.Trim();
        }

        private static string EscapeCsvField(string field)
        {
            if (field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }
            return field;
        }
    }
}
