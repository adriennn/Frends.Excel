using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Excel.Parse.Definitions;

public class Options
{
    /// <summary>
    /// Choose if exception should be thrown when conversion fails.
    /// </summary>
    [DefaultValue("true")]
    public bool ThrowErrorOnFailure { get; set; }
}
