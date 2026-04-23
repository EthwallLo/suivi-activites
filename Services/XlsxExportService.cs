using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using MonTableurApp.Models;

namespace MonTableurApp.Services
{
    public static class XlsxExportService
    {
        public static void ExportProjects(string filePath, IEnumerable<Projet> projets)
        {
            string[][] rows = BuildRows(projets).ToArray();

            using FileStream stream = File.Create(filePath);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

            WriteEntry(archive, "[Content_Types].xml", BuildContentTypesXml());
            WriteEntry(archive, "_rels/.rels", BuildRootRelsXml());
            WriteEntry(archive, "xl/workbook.xml", BuildWorkbookXml());
            WriteEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml());
            WriteEntry(archive, "xl/styles.xml", BuildStylesXml());
            WriteEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(rows));
        }

        private static IEnumerable<string[]> BuildRows(IEnumerable<Projet> projets)
        {
            yield return new[]
            {
                "Numéro projet",
                "Nom produit",
                "Client",
                "Demandeur",
                "Type d'activité",
                "Dossier racine",
                "Statut",
                "Date de début",
                "Date prévisionnelle",
                "Date de fin",
                "Référence produit",
                "Commentaires"
            };

            foreach (Projet projet in projets)
            {
                yield return new[]
                {
                    projet.NumeroProjet ?? string.Empty,
                    projet.NomProduit ?? string.Empty,
                    projet.Client ?? string.Empty,
                    projet.Demandeur ?? string.Empty,
                    projet.TypeActivite ?? string.Empty,
                    projet.DossierRacine ?? string.Empty,
                    projet.Statut ?? string.Empty,
                    projet.DateDebut ?? string.Empty,
                    projet.DatePrevisionnelle ?? string.Empty,
                    projet.DateFin ?? string.Empty,
                    projet.ReferenceProduit ?? string.Empty,
                    projet.Commentaires ?? string.Empty
                };
            }
        }

        private static string BuildContentTypesXml()
        {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                </Types>
                """;
        }

        private static string BuildRootRelsXml()
        {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """;
        }

        private static string BuildWorkbookXml()
        {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Projets" sheetId="1" r:id="rId1"/>
                  </sheets>
                </workbook>
                """;
        }

        private static string BuildWorkbookRelsXml()
        {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """;
        }

        private static string BuildStylesXml()
        {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="2">
                    <font>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                    </font>
                    <font>
                      <b/>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                    </font>
                  </fonts>
                  <fills count="2">
                    <fill><patternFill patternType="none"/></fill>
                    <fill><patternFill patternType="gray125"/></fill>
                  </fills>
                  <borders count="1">
                    <border>
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                  </borders>
                  <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                  </cellStyleXfs>
                  <cellXfs count="2">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
                  </cellXfs>
                  <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                  </cellStyles>
                </styleSheet>
                """;
        }

        private static string BuildWorksheetXml(IReadOnlyList<string[]> rows)
        {
            var builder = new StringBuilder();
            int lastRowIndex = rows.Count;
            string lastColumn = GetColumnName(rows[0].Length);

            builder.Append("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                """);
            builder.Append($"""<dimension ref="A1:{lastColumn}{lastRowIndex}"/>""");
            builder.Append("""
                <sheetViews>
                  <sheetView workbookViewId="0">
                    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
                  </sheetView>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15"/>
                <cols>
                  <col min="1" max="1" width="16" customWidth="1"/>
                  <col min="2" max="2" width="18" customWidth="1"/>
                  <col min="3" max="4" width="14" customWidth="1"/>
                  <col min="5" max="5" width="18" customWidth="1"/>
                  <col min="6" max="6" width="48" customWidth="1"/>
                  <col min="7" max="7" width="24" customWidth="1"/>
                  <col min="8" max="10" width="18" customWidth="1"/>
                  <col min="11" max="11" width="22" customWidth="1"/>
                  <col min="12" max="12" width="24" customWidth="1"/>
                </cols>
                <sheetData>
                """);

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                string[] row = rows[rowIndex];
                int excelRow = rowIndex + 1;
                builder.Append($"""<row r="{excelRow}">""");

                for (int columnIndex = 0; columnIndex < row.Length; columnIndex++)
                {
                    string cellRef = $"{GetColumnName(columnIndex + 1)}{excelRow}";
                    string style = rowIndex == 0 ? " s=\"1\"" : string.Empty;
                    builder.Append($"""<c r="{cellRef}" t="inlineStr"{style}><is><t xml:space="preserve">{EscapeXml(row[columnIndex])}</t></is></c>""");
                }

                builder.Append("</row>");
            }

            builder.Append("""
                </sheetData>
                <autoFilter ref="A1:L1"/>
                </worksheet>
                """);

            return builder.ToString();
        }

        private static void WriteEntry(ZipArchive archive, string entryName, string content)
        {
            ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using StreamWriter writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            writer.Write(content);
        }

        private static string GetColumnName(int index)
        {
            var builder = new StringBuilder();

            while (index > 0)
            {
                index--;
                builder.Insert(0, (char)('A' + (index % 26)));
                index /= 26;
            }

            return builder.ToString();
        }

        private static string EscapeXml(string value)
        {
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }
    }
}
