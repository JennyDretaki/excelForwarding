using System.ComponentModel;

namespace ExcelImporter.Models;

public sealed class ExcelRecord
{
    [DisplayName("Δικαστήριο")]
    public string Dikastirio { get; set; } = string.Empty;

    [DisplayName("Ονοματεπώνυμο")]
    public string Onomateponymo { get; set; } = string.Empty;

    [DisplayName("ΑΦΜ")]
    public string Afm { get; set; } = string.Empty;

    [DisplayName("Host Account")]
    public string HostAccount { get; set; } = string.Empty;

    [DisplayName("Σχόλια")]
    public string Scholia { get; set; } = string.Empty;

    [DisplayName("ΓΑΚ")]
    public string Gak { get; set; } = string.Empty;

    [DisplayName("ΕΑΚ")]
    public string Eak { get; set; } = string.Empty;

    [DisplayName("Ημερομηνία")]
    public string Imerominia { get; set; } = string.Empty;

    [Browsable(false)]
    public string SourceFilePath { get; set; } = string.Empty;
}
