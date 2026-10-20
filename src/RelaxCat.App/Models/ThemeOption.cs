namespace RelaxCat.App.Models;

public sealed class ThemeOption : ObservableModel
{
    private readonly UiText text;
    private readonly string key;
    public ThemeOption(UiText text, string key)
    {
        this.text = text; this.key = key;
        text.PropertyChanged += (_, _) => Notify(nameof(Label));
    }
    public string Label => text[key];
}
