using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using MonTableurApp.Models;

namespace MonTableurApp.Services
{
    public static class XlsxExportService
    {
        private static readonly CultureInfo DateCulture = CultureInfo.GetCultureInfo("fr-FR");
        private static readonly string[] DateFormats = { "dd/MM/yyyy", "d/M/yyyy" };

        public static void ExportProjects(string filePath, IEnumerable<Projet> projets)
        {
            ExportCell[][] rows = BuildRows(projets).ToArray();

            using FileStream stream = File.Create(filePath);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

            WriteEntry(archive, "[Content_Types].xml", BuildContentTypesXml());
            WriteEntry(archive, "_rels/.rels", BuildRootRelsXml());
            WriteEntry(archive, "xl/workbook.xml", BuildWorkbookXml());
            WriteEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml());
            WriteEntry(archive, "xl/styles.xml", BuildStylesXml());
            WriteEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(rows));
            WriteEntry(archive, "xl/worksheets/_rels/sheet1.xml.rels", BuildWorksheetRelsXml());
            WriteEntry(archive, "xl/tables/table1.xml", BuildTableXml(rows));
        }

        private static IEnumerable<ExportCell[]> BuildRows(IEnumerable<Projet> projets)
        {
            yield return new[]
            {
                ExportCell.Text("Numéro projet"),
                ExportCell.Text("Nom produit"),
                ExportCell.Text("Client"),
                ExportCell.Text("Demandeur"),
                ExportCell.Text("Type d'activité"),
                ExportCell.Text("Dossier racine"),
                ExportCell.Text("Statut"),
                ExportCell.Text("Date de début"),
                ExportCell.Text("Date prévisionnelle"),
                ExportCell.Text("Date de fin"),
                ExportCell.Text("Référence produit"),
                ExportCell.Text("Commentaires")
            };

            foreach (Projet projet in projets)
            {
                yield return new[]
                {
                    ExportCell.Text(projet.NumeroProjet),
                    ExportCell.Text(projet.NomProduit),
                    ExportCell.Text(projet.Client),
                    ExportCell.Text(projet.Demandeur),
                    ExportCell.Text(projet.TypeActivite),
                    ExportCell.Text(projet.DossierRacine),
                    ExportCell.Text(projet.Statut),
                    ExportCell.Date(projet.DateDebut),
                    ExportCell.Date(projet.DatePrevisionnelle),
                    ExportCell.Date(projet.DateFin),
                    ExportCell.Text(projet.ReferenceProduit),
                    ExportCell.Text(projet.Commentaires)
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
                  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
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

        private static string BuildWorksheetRelsXml()
        {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
                </Relationships>
                """;
        }

        private static string BuildStylesXml()
        {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <numFmts count="1">
                    <numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>
                  </numFmts>
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
                      <color rgb="FFFFFFFF"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                    </font>
                  </fonts>
                  <fills count="3">
                    <fill><patternFill patternType="none"/></fill>
                    <fill><patternFill patternType="gray125"/></fill>
                    <fill>
                      <patternFill patternType="solid">
                        <fgColor rgb="FF1F4E78"/>
                        <bgColor indexed="64"/>
                      </patternFill>
                    </fill>
                  </fills>
                  <borders count="2">
                    <border>
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left style="thin"><color rgb="FFD9E2EC"/></left>
                      <right style="thin"><color rgb="FFD9E2EC"/></right>
                      <top style="thin"><color rgb="FFD9E2EC"/></top>
                      <bottom style="thin"><color rgb="FFD9E2EC"/></bottom>
                      <diagonal/>
                    </border>
                  </borders>
                  <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                  </cellStyleXfs>
                  <cellXfs count="4">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1">
                      <alignment horizontal="center" vertical="center"/>
                    </xf>
                    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1">
                      <alignment horizontal="center" vertical="center"/>
                    </xf>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1">
                      <alignment vertical="center" wrapText="1"/>
                    </xf>
                  </cellXfs>
                  <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                  </cellStyles>
                </styleSheet>
                """;
        }

        private static string BuildWorksheetXml(IReadOnlyList<ExportCell[]> rows)
        {
            var builder = new StringBuilder();
            int lastRowIndex = rows.Count;
            string lastColumn = GetColumnName(rows[0].Length);

            builder.Append("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                """);
            builder.Append($"""<dimension ref="A1:{lastColumn}{lastRowIndex}"/>""");
            builder.Append("""
                <sheetViews>
                  <sheetView workbookViewId="0">
                    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
                    <selection pane="bottomLeft"/>
                  </sheetView>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="18"/>
                <cols>
                  <col min="1" max="1" width="16" customWidth="1"/>
                  <col min="2" max="2" width="22" customWidth="1"/>
                  <col min="3" max="4" width="16" customWidth="1"/>
                  <col min="5" max="5" width="20" customWidth="1"/>
                  <col min="6" max="6" width="56" customWidth="1"/>
                  <col min="7" max="7" width="26" customWidth="1"/>
                  <col min="8" max="10" width="18" customWidth="1"/>
                  <col min="11" max="11" width="22" customWidth="1"/>
                  <col min="12" max="12" width="34" customWidth="1"/>
                </cols>
                <sheetData>
                """);

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                ExportCell[] row = rows[rowIndex];
                int excelRow = rowIndex + 1;
                string rowHeight = rowIndex == 0 ? " ht=\"24\" customHeight=\"1\"" : string.Empty;
                builder.Append($"""<row r="{excelRow}"{rowHeight}>""");

                for (int columnIndex = 0; columnIndex < row.Length; columnIndex++)
                {
                    string cellRef = $"{GetColumnName(columnIndex + 1)}{excelRow}";
                    builder.Append(BuildCellXml(cellRef, row[columnIndex], rowIndex == 0));
                }

                builder.Append("</row>");
            }

            builder.Append($"""
                </sheetData>
                <tableParts count="1">
                  <tablePart r:id="rId1"/>
                </tableParts>
                </worksheet>
                """);

            return builder.ToString();
        }

        private static string BuildTableXml(IReadOnlyList<ExportCell[]> rows)
        {
            string lastColumn = GetColumnName(rows[0].Length);
            string tableRef = $"A1:{lastColumn}{rows.Count}";
            var builder = new StringBuilder();

            builder.Append($"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       id="1"
                       name="SuiviProjets"
                       displayName="SuiviProjets"
                       ref="{tableRef}"
                       totalsRowShown="0">
                  <autoFilter ref="{tableRef}"/>
                  <tableColumns count="{rows[0].Length}">
                """);

            for (int columnIndex = 0; columnIndex < rows[0].Length; columnIndex++)
            {
                builder.Append($"""<tableColumn id="{columnIndex + 1}" name="{EscapeAttribute(rows[0][columnIndex].TextValue)}"/>""");
            }

            builder.Append("""
                  </tableColumns>
                  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
                """);

            return builder.ToString();
        }

        private static string BuildCellXml(string cellRef, ExportCell cell, bool isHeader)
        {
            if (!isHeader && cell.DateValue is DateTime dateValue)
            {
                string dateSerial = dateValue.ToOADate().ToString(CultureInfo.InvariantCulture);
                return $"""<c r="{cellRef}" s="2"><v>{dateSerial}</v></c>""";
            }

            string style = isHeader ? "1" : "3";
            return $"""<c r="{cellRef}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{EscapeXml(cell.TextValue)}</t></is></c>""";
        }

        private static DateTime? ParseDate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            string trimmedValue = value.Trim();

            if (DateTime.TryParseExact(trimmedValue, DateFormats, DateCulture, DateTimeStyles.None, out DateTime exactDate))
            {
                return exactDate.Date;
            }

            return DateTime.TryParse(trimmedValue, DateCulture, DateTimeStyles.None, out DateTime parsedDate)
                ? parsedDate.Date
                : null;
        }

        private static string CleanXmlText(string? value)
        {
            string input = value ?? string.Empty;
            var builder = new StringBuilder(input.Length);

            foreach (char current in input)
            {
                if (XmlConvert.IsXmlChar(current))
                {
                    builder.Append(current);
                }
            }

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

        private static string EscapeXml(string? value)
        {
            return CleanXmlText(value)
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        private static string EscapeAttribute(string? value)
        {
            return EscapeXml(value)
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        private sealed class ExportCell
        {
            private ExportCell(string textValue, DateTime? dateValue)
            {
                TextValue = CleanXmlText(textValue);
                DateValue = dateValue;
            }

            public string TextValue { get; }

            public DateTime? DateValue { get; }

            public static ExportCell Text(string? value)
            {
                return new ExportCell(value ?? string.Empty, null);
            }

            public static ExportCell Date(string? value)
            {
                return new ExportCell(value ?? string.Empty, ParseDate(value));
            }
        }
    }
}
