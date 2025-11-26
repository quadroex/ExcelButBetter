using CommunityToolkit.Maui.Views;

namespace ExcelButBetter;

public enum UserChoice
{
    Cancel,
    Save,
    Discard
}

public partial class UnsavedChangesPopup : Popup
{
    public UserChoice Result { get; private set; } = UserChoice.Cancel;

    public UnsavedChangesPopup() { InitializeComponent(); }

    private async void OnSaveClicked(object sender, EventArgs e)
    {
        Result = UserChoice.Save;
        await CloseAsync();
    }

    private async void OnDontSaveClicked(object sender, EventArgs e)
    {
        Result = UserChoice.Discard;
        await CloseAsync();
    }

    private async void OnCancelClicked(object sender, EventArgs e)
    {
        Result = UserChoice.Cancel;
        await CloseAsync();
    }
}