using System.IO.Compression;
using System.Text;

namespace RelaxCat.Core.Tests;

internal sealed record SheetSpec(string Name, string Xml, bool Hidden = false);

internal static class OoxmlWorkbookFixture
{
    internal const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal static string Create(string path, string stylesXml, params SheetSpec[] sheets)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            Add(archive, "[Content_Types].xml", ContentTypes(sheets.Length));
            Add(archive, "_rels/.rels", RootRelationships());
            Add(archive, "xl/workbook.xml", Workbook(sheets));
            Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships(sheets.Length));
            Add(archive, "xl/styles.xml", stylesXml);
            Add(archive, "xl/theme/theme1.xml", ThemeXml());
            for (var index = 0; index < sheets.Length; index++)
                Add(archive, $"xl/worksheets/sheet{index + 1}.xml", sheets[index].Xml);
        }
        return path;
    }

    internal static void AddPart(string path, string partName, byte[] bytes)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.CreateEntry(partName);
        using var stream = entry.Open();
        stream.Write(bytes);
    }

    internal static byte[] ReadPart(string path, string partName)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry(partName) ?? throw new InvalidOperationException("Missing part " + partName);
        using var stream = entry.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    internal static string[] PartNames(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        return archive.Entries.Select(entry => entry.FullName).ToArray();
    }

    internal static string CellXfs(params string[] formats) =>
        $"<cellXfs count=\"{formats.Length}\">{string.Concat(formats)}</cellXfs>";

    internal static string StylesXml(string cellXfs, string? fonts = null, string? fills = null,
        string? borders = null, string? numberFormats = null)
    {
        fonts ??= """
                  <fonts count="1">
                    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
                  </fonts>
                  """;
        fills ??= """
                  <fills count="2">
                    <fill><patternFill patternType="none"/></fill>
                    <fill><patternFill patternType="gray125"/></fill>
                  </fills>
                  """;
        borders ??= """
                    <borders count="1">
                      <border><left/><right/><top/><bottom/><diagonal/></border>
                    </borders>
                    """;
        return $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="{MainNamespace}">
                {numberFormats ?? ""}
                {fonts}
                {fills}
                {borders}
                <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                {cellXfs}
                <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
                <dxfs count="0"/>
                <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
                </styleSheet>
                """;
    }

    internal static string Worksheet(string dimension, string sheetData) => $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="{MainNamespace}">
          <dimension ref="{dimension}"/>
          <sheetViews><sheetView workbookViewId="0"/></sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <sheetData>{sheetData}</sheetData>
          <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        </worksheet>
        """;

    internal static string ThemeXml(string accent1 = "4472C4")
    {
        static string Srgb(string value) => $"""<a:srgbClr val="{value}"/>""";
        return $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
                  <a:themeElements>
                    <a:clrScheme name="Office">
                      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
                      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
                      <a:dk2>{Srgb("44546A")}</a:dk2>
                      <a:lt2>{Srgb("E7E6E6")}</a:lt2>
                      <a:accent1>{Srgb(accent1)}</a:accent1>
                      <a:accent2>{Srgb("ED7D31")}</a:accent2>
                      <a:accent3>{Srgb("A5A5A5")}</a:accent3>
                      <a:accent4>{Srgb("FFC000")}</a:accent4>
                      <a:accent5>{Srgb("5B9BD5")}</a:accent5>
                      <a:accent6>{Srgb("70AD47")}</a:accent6>
                      <a:hlink>{Srgb("0563C1")}</a:hlink>
                      <a:folHlink>{Srgb("954F72")}</a:folHlink>
                    </a:clrScheme>
                  </a:themeElements>
                </a:theme>
                """;
    }

    private static string ContentTypes(int sheetCount)
    {
        var sheets = new StringBuilder();
        for (var index = 1; index <= sheetCount; index++)
            sheets.AppendLine(
                $"  <Override PartName=\"/xl/worksheets/sheet{index}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        return $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                {sheets}</Types>
                """;
    }

    private static string RootRelationships() => $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="{PackageRelationshipNamespace}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private static string Workbook(SheetSpec[] sheets)
    {
        var entries = new StringBuilder();
        for (var index = 0; index < sheets.Length; index++)
        {
            var state = sheets[index].Hidden ? " state=\"hidden\"" : "";
            entries.AppendLine(
                $"    <sheet name=\"{sheets[index].Name}\" sheetId=\"{index + 1}\"{state} r:id=\"rId{index + 1}\"/>");
        }
        return $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="{MainNamespace}" xmlns:r="{RelationshipNamespace}">
                  <sheets>
                {entries}  </sheets>
                </workbook>
                """;
    }

    private static string WorkbookRelationships(int sheetCount)
    {
        var entries = new StringBuilder();
        for (var index = 1; index <= sheetCount; index++)
            entries.AppendLine(
                $"  <Relationship Id=\"rId{index}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{index}.xml\"/>");
        entries.AppendLine(
            $"  <Relationship Id=\"rId{sheetCount + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        entries.AppendLine(
            $"  <Relationship Id=\"rId{sheetCount + 2}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>");
        return $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{PackageRelationshipNamespace}">
                {entries}</Relationships>
                """;
    }

    private static void Add(ZipArchive archive, string name, string contents)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(contents);
    }
}
