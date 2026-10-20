using System.Globalization;
using System.Text.RegularExpressions;

namespace RelaxCat.Core.Imaging;

public readonly record struct ScaleRuleContext(double Width, double Height, double DpiX, double DpiY)
{
    public double Dpi => (DpiX + DpiY) / 2d;
    public double AspectRatio => Height == 0 ? 0 : Width / Height;

    public double Get(string name) => name.ToLowerInvariant() switch
    {
        "width" => Width,
        "height" => Height,
        "dpi" => Dpi,
        "dpix" => DpiX,
        "dpiy" => DpiY,
        "aspectratio" => AspectRatio,
        _ => throw new ScaleRuleException($"Unknown variable '{name}'.")
    };
}

public sealed class ScaleRuleException : Exception
{
    public ScaleRuleException(string message) : base(message) { }
}

public sealed class ScaleRuleEngine
{
    private static readonly Regex ScaleRegex = new(@"^scale\s*=\s*(?<value>[0-9]+(?:\.[0-9]+)?)\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ComparisonRegex = new(@"^(?<name>[A-Za-z][A-Za-z0-9]*)\s*(?<op><=|>=|==|!=|<|>)\s*(?<value>-?[0-9]+(?:\.[0-9]+)?)$", RegexOptions.Compiled);

    public double? Evaluate(string rules, ScaleRuleContext context)
    {
        foreach (var rule in ParseRules(rules))
            if (EvaluateCondition(rule.Condition, context))
                return rule.Scale;
        return null;
    }

    public IReadOnlyList<ScaleRule> ParseRules(string text)
    {
        var lines = text.Replace("\r", string.Empty).Split('\n');
        var result = new List<ScaleRule>();
        string? condition = null;

        foreach (var raw in lines)
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('#')) continue;

            var scaleMatch = ScaleRegex.Match(line);
            if (scaleMatch.Success)
            {
                if (condition is null) throw new ScaleRuleException("Scale assignment has no condition.");
                var scale = double.Parse(scaleMatch.Groups["value"].Value, CultureInfo.InvariantCulture);
                if (scale <= 0) throw new ScaleRuleException("Scale must be positive.");
                result.Add(new ScaleRule(condition, scale));
                condition = null;
                continue;
            }

            if (condition is not null) throw new ScaleRuleException($"Expected scale assignment after '{condition}'.");
            condition = line;
        }

        if (condition is not null) throw new ScaleRuleException($"Missing scale assignment after '{condition}'.");
        return result;
    }

    private static bool EvaluateCondition(string expression, ScaleRuleContext context)
    {
        // Initial grammar: comparisons > and > or. Parentheses are intentionally deferred.
        var orParts = Regex.Split(expression, @"\s+or\s+", RegexOptions.IgnoreCase);
        return orParts.Any(orPart => Regex.Split(orPart.Trim(), @"\s+and\s+", RegexOptions.IgnoreCase)
            .All(part => EvaluateComparison(part.Trim(), context)));
    }

    private static bool EvaluateComparison(string expression, ScaleRuleContext context)
    {
        var match = ComparisonRegex.Match(expression);
        if (!match.Success) throw new ScaleRuleException($"Invalid comparison '{expression}'.");

        var left = context.Get(match.Groups["name"].Value);
        var right = double.Parse(match.Groups["value"].Value, CultureInfo.InvariantCulture);
        return match.Groups["op"].Value switch
        {
            "<" => left < right,
            "<=" => left <= right,
            ">" => left > right,
            ">=" => left >= right,
            "==" => Math.Abs(left - right) < 1e-9,
            "!=" => Math.Abs(left - right) >= 1e-9,
            _ => false
        };
    }
}

public sealed record ScaleRule(string Condition, double Scale);
