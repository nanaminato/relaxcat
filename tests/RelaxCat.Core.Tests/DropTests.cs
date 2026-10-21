using RelaxCat.Core.Import;

namespace RelaxCat.Core.Tests;

public sealed class DropPathResolverTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "relaxcat-drop-" + Guid.NewGuid().ToString("N"));

    public DropPathResolverTests() => Directory.CreateDirectory(root);
    public void Dispose() { try { Directory.Delete(root, true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }

    private string Create(string relative)
    {
        var path = Path.Combine(root, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "test");
        return path;
    }

    [Fact]
    public void Folder_drop_collects_images_from_every_subfolder_in_name_order()
    {
        Create("b.png");
        Create("a.png");
        Create("nested/inner.JPG");
        Create("nested/deep/leaf.bmp");
        Create("nested/deep/notes.txt");
        var result = DropPathResolver.Resolve(new[] { root });
        Assert.Equal(new[] { "a.png", "b.png", "inner.JPG", "leaf.bmp" }, result.Images.Select(Path.GetFileName));
        Assert.Empty(result.Issues);
        Assert.Equal(new[] { "notes.txt" }, result.Ignored.Select(Path.GetFileName));
    }

    [Fact]
    public void Duplicated_drops_import_each_image_once()
    {
        var image = Create("shared/a.png");
        Create("shared/nested/a.png");
        var result = DropPathResolver.Resolve(new[] { root, image, Path.Combine(root, "shared") + Path.DirectorySeparatorChar });
        Assert.Equal(new[] { Path.Combine("shared", "a.png"), Path.Combine("shared", "nested", "a.png") },
            result.Images.Select(path => Path.GetRelativePath(root, path)));
    }

    [Fact]
    public void Supported_formats_are_matched_case_insensitively()
    {
        var paths = new[] { "one.PNG", "two.Jpeg", "three.BmP", "four.gif" }.Select(Create).ToArray();
        var result = DropPathResolver.Resolve(new[] { Path.Combine(root, "one.PNG"), Path.Combine(root, "two.Jpeg"), Path.Combine(root, "three.BmP"), Path.Combine(root, "four.gif") });
        Assert.Equal(paths.Take(3), result.Images);
        Assert.Equal(new[] { paths[3] }, result.Ignored);
    }

    [Fact]
    public void Missing_paths_and_image_free_folders_are_reported_without_failing_the_drop()
    {
        Create("plain/readme.txt");
        var missing = Path.Combine(root, "gone");
        var result = DropPathResolver.Resolve(new[] { Path.Combine(root, "plain"), missing });
        Assert.Empty(result.Images);
        Assert.Equal(new[] { new DropIssue(missing, DropIssueKind.NotFound) }, result.Issues);
    }
}
