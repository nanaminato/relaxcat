namespace RelaxCat.App.Models;

public sealed class LocalizedOption : ObservableModel
{
    private readonly UiText text;
    private readonly string key;
    public LocalizedOption(UiText text, string key)
    {
        this.text = text; this.key = key;
        text.PropertyChanged += (_, _) => Notify(nameof(Label));
    }
    public string Label => text[key];
}
