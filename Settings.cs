internal sealed class Settings : ApplicationSettingsBase, INotifyPropertyChanged
{
    // Add the missing property
    public bool IsConverterDetached { get; set; }
}