using System.Text;
using System.Text.RegularExpressions;

namespace RelaxCat.Core.Excel;

/// <summary>Edits lexical XML spans so retained formats and surrounding bytes never get serialized.</summary>
internal static class CellXfsByteRewriter
{
    internal static byte[] Rewrite(byte[] original, int[] map, string table = "cellXfs", string item = "xf")
    {
        var encoding = DetectEncoding(original);
        var text = encoding.GetString(original);
        var edits = new List<(int Start, int Length, string Replacement)>();
        var depth = 0;
        var tableDepth = -1;
        var formatStart = -1;
        var index = 0;
        var seen = new HashSet<int>();
        // Quoted attribute values may contain '>'; comments/CDATA/PIs may contain fake tags.
        foreach (Match token in Regex.Matches(text,
                     "<!--.*?-->|<!\\[CDATA\\[.*?\\]\\]>|<\\?.*?\\?>|</?[^\\s/>]+(?:[^>\"']|\"[^\"]*\"|'[^']*')*>",
                     RegexOptions.Singleline))
        {
            var tag = token.Value;
            if (tag.StartsWith("<!", StringComparison.Ordinal) || tag.StartsWith("<?", StringComparison.Ordinal)) continue;
            var closing = tag.StartsWith("</", StringComparison.Ordinal);
            var empty = tag.EndsWith("/>", StringComparison.Ordinal);
            var name = Regex.Match(tag, "^</?([^\\s/>]+)").Groups[1].Value.Split(':')[^1];
            if (closing) depth--;
            if (!closing && depth == 1 && name == table)
            {
                tableDepth = depth;
                var count = Regex.Match(tag, "(?<=\\s)count\\s*=\\s*(['\"])[^'\"]*\\1");
                var replacement = "count=\"" + map.Distinct().Count().ToString(System.Globalization.CultureInfo.InvariantCulture) + "\"";
                if (count.Success) edits.Add((token.Index + count.Index, count.Length, replacement));
                else edits.Add((token.Index + tag.Length - (empty ? 2 : 1), 0, " " + replacement));
            }
            else if (tableDepth >= 0 && !closing && depth == tableDepth + 1 && name == item)
                formatStart = token.Index;

            if (tableDepth >= 0 && name == item && depth == tableDepth + 1 && (closing || empty))
            {
                if (!seen.Add(map[index++]))
                    edits.Add((formatStart, token.Index + token.Length - formatStart, ""));
            }
            if (closing && depth == tableDepth) break;
            if (!closing && !empty) depth++;
        }
        if (index != map.Length) throw new InvalidDataException("Cannot locate all cell formats in the original XML.");
        using var output = new MemoryStream();
        var position = 0;
        var characterPosition = 0;
        foreach (var edit in edits.OrderBy(edit => edit.Start))
        {
            var start = position + encoding.GetByteCount(text.AsSpan(characterPosition, edit.Start - characterPosition));
            var end = start + encoding.GetByteCount(text.AsSpan(edit.Start, edit.Length));
            output.Write(original.AsSpan(position, start - position));
            output.Write(encoding.GetBytes(edit.Replacement));
            position = end;
            characterPosition = edit.Start + edit.Length;
        }
        output.Write(original.AsSpan(position));
        return output.ToArray();
    }
    private static Encoding DetectEncoding(byte[] original)
    {
        // XML permits UTF-8 and UTF-16 without a BOM; honor an explicit encoding as well.
        Encoding encoding;
        if (original.AsSpan().StartsWith(new byte[] { 0xff, 0xfe }) ||
            (original.Length > 1 && original[0] == '<' && original[1] == 0))
            encoding = Encoding.Unicode;
        else if (original.AsSpan().StartsWith(new byte[] { 0xfe, 0xff }) ||
                 (original.Length > 1 && original[0] == 0 && original[1] == '<'))
            encoding = Encoding.BigEndianUnicode;
        else
        {
            var declaration = Regex.Match(Encoding.ASCII.GetString(original, 0, Math.Min(original.Length, 256)),
                "^<\\?xml\\s+[^?]*encoding\\s*=\\s*['\"]([^'\"]+)['\"]");
            encoding = declaration.Success ? Encoding.GetEncoding(declaration.Groups[1].Value) : Encoding.UTF8;
        }
        return encoding;
    }

    internal static byte[] RemapBaseReferences(byte[] original, int[] map)
    {
        var encoding = DetectEncoding(original);
        var text = encoding.GetString(original);
        var parents = new Stack<string>();
        var result = Regex.Replace(text,
            "<!--.*?-->|<!\\[CDATA\\[.*?\\]\\]>|<\\?.*?\\?>|</?[^\\s/>]+(?:[^>\"']|\"[^\"]*\"|'[^']*')*>", token =>
        {
            var tag = token.Value;
            if (tag.StartsWith("<!") || tag.StartsWith("<?")) return tag;
            var closing = tag.StartsWith("</");
            var empty = tag.EndsWith("/>");
            var name = Regex.Match(tag, "^</?([^\\s/>]+)").Groups[1].Value.Split(':')[^1];
            if (closing) { parents.Pop(); return tag; }
            if (parents.Count == 2 && ((parents.Peek() == "cellXfs" && name == "xf") ||
                                      (parents.Peek() == "cellStyles" && name == "cellStyle")))
                tag = Regex.Replace(tag, "(?<=\\s)xfId\\s*=\\s*(['\"])([^'\"]*)\\1", attribute =>
                {
                    var id = int.Parse(attribute.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
                    if (id >= map.Length) throw new InvalidDataException("Invalid base style reference.");
                    return id == map[id] ? attribute.Value : "xfId=" + attribute.Groups[1].Value +
                        map[id].ToString(System.Globalization.CultureInfo.InvariantCulture) + attribute.Groups[1].Value;
                });
            if (!empty) parents.Push(name);
            return tag;
        }, RegexOptions.Singleline);
        return encoding.GetBytes(result);
    }

}
