using System.IO.Compression;
using System.Xml.Linq;
using RelaxCat.Core.Excel;

namespace RelaxCat.App.SmokeTests;

internal static class NamedStyleFixtureChecks
{
    internal static void Run()
    {
        var input = Path.GetFullPath("artifacts/named-style-fixture-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var output = Path.ChangeExtension(input, ".optimized.xlsx");
        try
        {
            ExcelToolChecks.CreateNamedStyleWorkbook(input);
            CheckNamedIdentities(input, 3);
            var options = new ExcelStyleOptimizationOptions(RemoveCopiedNamedStyles: true);
            var analysis = new ExcelStyleAnalyzer().Analyze(input, options: options);
            if (!analysis.NamedStyleCleanupCandidates.SequenceEqual(["Custom_1"]))
                throw new Exception("Unreferenced copied-name fixture was not eligible for explicit cleanup");
            var result = new ExcelStyleOptimizer().Optimize(input, output, options: options);
            if (!result.Verified) throw new Exception("Copied-name fixture optimization was not verified");
            CheckNamedIdentities(output, 2);

            ExcelToolChecks.CreateNamedStyleWorkbook(input, referenceCandidateFromUnusedFormat: true);
            CheckNamedIdentities(input, 3);
            foreach (var gc in new[] { false, true })
            {
                analysis = new ExcelStyleAnalyzer().Analyze(input,
                    options: new ExcelStyleOptimizationOptions(CollectUnusedStyles: gc, RemoveCopiedNamedStyles: true));
                var expectedCandidates = gc ? 1 : 0;
                if (analysis.NamedStyleCleanupCandidates.Count != expectedCandidates)
                    throw new Exception("Copied-name cleanup did not respect retained cellXf references after optional GC");
            }
            Console.WriteLine("PASS: App named-style fixtures use distinct identities and protect retained cellXf associations after optional GC.");
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    private static void CheckNamedIdentities(string path, int expectedCount)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry("xl/styles.xml")!.Open();
        var styles = XDocument.Load(stream);
        var names = styles.Root!.Elements().Single(element => element.Name.LocalName == "cellStyles").Elements().ToArray();
        if (names.Length != expectedCount ||
            names.Select(element => (string?)element.Attribute("name")).Distinct(StringComparer.OrdinalIgnoreCase).Count() != expectedCount ||
            names.Select(element => (int)element.Attribute("xfId")!).Distinct().Count() != expectedCount)
            throw new Exception("App fixture or output has duplicate named-style names or base XF identities");
    }
}
