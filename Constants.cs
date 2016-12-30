namespace PuzXls
{
  public class Constants
  {
    //  0 - H   - ushort - overall checksum
    //  1 - 11s - string - file magic "ACROSS&DOWN"
    //    - x   - null term
    //  2 - H   - ushort - chksum header
    //  3 - Q   - ulong  - chksum magic
    //  4 - 3s  - string - file version
    //    - x   - null term
    //  5 - 2s  - string - <unknown> ( no terminator)
    //  6 - H   - ushort - scrambled checksum
    //  7 - 12s - string - <unknown>
    //  8 - B   - int    - width
    //  9 - B   - int    - height
    // 10 - H   - ushort - clue count
    // 11 - H   - ushort - puzzle type bitmask
    // 12 - H   - ushort - solution state
    public const string HEADER_FORMAT = @"<
             H 11s         xH
             Q       3s x2s H
             12s         BBH
             H H ";

    //   - <   -        - little endian
    // 0 - 4s  - string - code
    // 1 - H   - short  - length of extension name
    // 2 - H   - short  - checksum
    public const string EXTENSION_HEADER_FORMAT = "< 4s  H H ";

    public const string HEADER_CKSUM_FORMAT     = "<BBH H H ";
    public const string MASKSTRING              = "ICHEATED";
    public const string DEFAULT_ENCODING        = "us-ascii";
    public const string ACROSSDOWN              = "ACROSS&DOWN";
    public const string BLACKSQUARE             = ".";
  }

  public enum PuzzleType
  {
    Normal      = 0x0001,
    Diagramless = 0x0401
  }

  // the following diverges from the documentation
  // but works for the files I've tested
  public enum SolutionState
  {
    Unlocked = 0x0000, // solution is available in plaintext
    Locked   = 0x0004  // solution is locked (scrambled) with a key
  }

  public enum GridMarkup
  {
    Default             = 0x00, // ordinary grid cell
    PreviouslyIncorrect = 0x10, // marked incorrect at some point
    Incorrect           = 0x20, // currently showing incorrect
    Revealed            = 0x40, // user got a hint
    Circled             = 0x80  // circled
  }
}