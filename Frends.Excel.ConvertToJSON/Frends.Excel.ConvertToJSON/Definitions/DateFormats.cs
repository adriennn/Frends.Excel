namespace Frends.Excel.ConvertToJSON.Definitions;

/// <summary>
/// Date formats enum.
/// </summary>
public enum DateFormats
{
    /// <summary>
    /// Default format, will use `CurrentCulture.DateTimeFormat.ShortDatePattern`
    /// for `ShortDatePattern` option is set to `true`, otherwise will use 
    /// `CurrentCulture.DateTimeFormat.DateTimeFormat`.
    /// </summary>
    DEFAULT,

    /// <summary>
    /// Will use `dd/MM/yyyy` if `ShortDatePattern` option is set to `true`,
    /// otherwise will use `dd/MM/yyyy H:mm:ss`.
    /// </summary>
    DDMMYYYY,
    /// <summary>
    /// Will use `MM/dd/yyyy` if `ShortDatePattern` option is set to `true`,
    /// otherwise will use `MM/dd/yyyy h:mm:ss tt`.
    /// </summary>
    MMDDYYYY,
    
    /// <summary>
    /// Will use `yyyy/MM/dd` if `ShortDatePattern` option is set to `true`,
    /// otherwise will use `yyyy/MM/dd H:mm:ss`.
    /// </summary>
    YYYYMMDD
}
