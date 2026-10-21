namespace RelaxCat.Core.Import;

/// <summary>Image formats accepted by manual import, drag-and-drop and folder monitoring.</summary>
public static class ImageFileFormats
{
    public static bool Supports(string path) =>
        Path.GetExtension(path).ToLowerInvariant() is ".png" or ".jpg" or ".jpeg" or ".bmp";
}

public enum DropIssueKind
{
    /// <summary>The dropped path no longer exists (deleted, moved or mistyped).</summary>
    NotFound,
    /// <summary>The dropped path exists but its contents could not be enumerated.</summary>
    Unreadable
}

public sealed record DropIssue(string Path, DropIssueKind Kind);

/// <summary>
/// Expands dropped files and folders into importable image paths. Folders are walked recursively so a single
/// dropped folder contributes the images of its whole subtree; repeated paths, cycle-forming links and
/// unreadable subtrees do not abort the rest of the drop.
/// </summary>
public static class DropPathResolver
{
    public sealed record Resolution(IReadOnlyList<string> Images, IReadOnlyList<string> Ignored, IReadOnlyList<DropIssue> Issues);

    public static Resolution Resolve(IEnumerable<string> droppedPaths)
    {
        ArgumentNullException.ThrowIfNull(droppedPaths);
        var images = new List<string>();
        var ignored = new List<string>();
        var issues = new List<DropIssue>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var dropped in droppedPaths)
        {
            if (string.IsNullOrWhiteSpace(dropped)) continue;
            string full;
            try { full = Path.GetFullPath(dropped); }
            catch (Exception error) when (error is ArgumentException or NotSupportedException or PathTooLongException)
            {
                issues.Add(new DropIssue(dropped, DropIssueKind.NotFound)); continue;
            }
            if (File.Exists(full))
            {
                try { AddFile(full, images, ignored, seen); }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException) { issues.Add(new DropIssue(full, DropIssueKind.Unreadable)); }
            }
            else if (Directory.Exists(full)) Collect(full, images, ignored, issues, seen);
            else issues.Add(new DropIssue(full, DropIssueKind.NotFound));
        }
        return new Resolution(images, ignored, issues);
    }

    private static void Collect(string root, List<string> images, List<string> ignored, List<DropIssue> issues, HashSet<string> seen)
    {
        // Explicit stack: deep trees must not overflow the call stack, and reverse pushes keep name order.
        var pending = new Stack<string>();
        pending.Push(root);
        while (pending.Count > 0)
        {
            var folder = pending.Pop();
            string[] files;
            try { files = Directory.GetFiles(folder); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { issues.Add(new DropIssue(folder, DropIssueKind.Unreadable)); continue; }
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);
            foreach (var file in files)
            {
                try { AddFile(file, images, ignored, seen); }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException) { issues.Add(new DropIssue(file, DropIssueKind.Unreadable)); }
            }
            string[] children;
            try { children = Directory.GetDirectories(folder); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { issues.Add(new DropIssue(folder, DropIssueKind.Unreadable)); continue; }
            Array.Sort(children, StringComparer.OrdinalIgnoreCase);
            for (var i = children.Length - 1; i >= 0; i--)
            {
                if (IsLink(children[i])) continue; // Junctions and symlinks could loop back into the same subtree.
                pending.Push(children[i]);
            }
        }
    }

    private static bool IsLink(string folder)
    {
        try { return (File.GetAttributes(folder) & FileAttributes.ReparsePoint) != 0; }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { return true; }
    }

    private static void AddFile(string path, List<string> images, List<string> ignored, HashSet<string> seen)
    {
        if (!ImageFileFormats.Supports(path)) { ignored.Add(path); return; }
        if (seen.Add(path)) images.Add(path);
    }
}
