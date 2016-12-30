using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

using ClosedXML.Excel;

namespace PuzXls
{
  class Program
  {
    static void Main(string[] args)
    {
      var p = Puz.Read(args.Length == 0? Console.ReadLine() : args[0]);

      var workbook  = new XLWorkbook();
      var worksheet = workbook.Worksheets.Add("Puz");

      worksheet.Rows().Height   = 20;
      worksheet.Columns().Width = 4;

      var grid = p.Grid;

      int height = grid.GetLength(0);

      for (int row = 0; row < height; row++)
      for (int col = 0; col < grid.GetLength(1); col++)
      {
        string cellValue = grid[row, col];
        var    cell      = worksheet.Cell(row + 1, col + 1);

        if (cellValue == ".")
          cell.Style.Fill.BackgroundColor = XLColor.WhiteSmoke;

        if (cellValue == "-" || cellValue == ".")
          cellValue = String.Empty;

        cell.Value = cellValue;

        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        cell.Style.Alignment.Vertical   = XLAlignmentVerticalValues.Center;
      }

      var clueNumbering = new DefaultClueNumbering(p.Fill, p.Clues, p.Width, p.Height);

      int clueRowIndex = 1;

      foreach (var acrossClue in clueNumbering.Across)
        worksheet.Cell(clueRowIndex++, grid.GetLength(1) + 2).Value =  $"{acrossClue.Number} ({worksheet.Cell(acrossClue.Row + 1, acrossClue.Col + 1).Address.ToString()}) {acrossClue.Direction}: {acrossClue.Clue}";

      foreach (var downClue in clueNumbering.Down)
       worksheet.Cell(clueRowIndex++, grid.GetLength(1) + 2).Value =  $"{downClue.Number} ({worksheet.Cell(downClue.Row + 1, downClue.Col + 1).Address.ToString()}) {downClue.Direction}: {downClue.Clue}";

      workbook.SaveAs(@"C:\test.xlsx");
    }
  }

  // refer to Extensions as Extensions.Rebus, Extensions.Markup
  public static class Extensions
  {
    // grid of rebus indices: 0 for non-rebus;
    // i+1 for key i into RebusSolutions map
    public static readonly string Rebus ="GRBS";

    // map of rebus solution entries eg 0:HEART;1:DIAMOND;17:CLUB;23:SPADE;
    public static readonly string RebusSolutions = "RTBL";

    // user's rebus entries
    public static readonly string RebusFill = "RUSR";

    // timer state: 'a,b' where a is the number of seconds elapsed and
    // b is a boolean (0,1) for whether the timer is running
    public static readonly string Timer = "LTIM";

    // grid cell markup: previously incorrect: 0x10;
    // currently incorrect: 0x20,
    // hinted: 0x40,
    // circled: 0x80
    public static readonly string Markup = "GEXT";
  }

  public class Puz
  {
    /// <summary>
    /// Read a .puz file and return the Puzzle object.
    /// throws PuzzleFormatError if there's any problem with the file format.
    /// </summary>
    public static Puz Read(string filePath)
    {
      using (var fileStream   = new FileStream(filePath, FileMode.Open))
      using (var binaryReader = new BinaryReader(fileStream))
      {
        byte[] data = binaryReader.ReadBytes((int)fileStream.Length);

        var puz = new Puz();

        puz.Load(data);

        return puz;
      }
    }

    private string        _postscript;
    private string        _title;
    private string        _author;
    private string        _copyright;

    private byte          _width;
    private byte          _height;

    private string        _version;
    private string        _fileVersion;

    private string        _unk1;
    private string        _unk2;

    private ushort        _scrambledChecksum;
    private string        _fill;
    private string        _solution;
    private List<string>  _clues;
    private string        _notes;
    private Dictionary<string, string> _extensions;

    private PuzzleType    _puzzleType;
    private SolutionState _solutionState;
    private List<string>  _helpers;

    public string Fill => _fill;
    public int    Width => (int)_width;
    public int    Height => (int)_height;

    public List<string> Clues => _clues;
    public string[,]    Grid
    {
      get
      {
        var grid = new string[_width, _height];

        for (int x = 0; x < _width;  x++)
        for (int y = 0; y < _height; y++)
          grid[x,y] = _fill[(x * _height) + y].ToString();

        return grid;
      }
    }

    public Puz()
    {
      // Initializes a blank puzzle
      _postscript        = "";
      _title             = "";
      _author            = "";
      _copyright         = "";
      _width             = 0;
      _height            = 0;
      _version           = "1.3";
      _fileVersion       = "1.3\0";  // default

      // these are bytes that might be unused
      _unk1              = "\0\0";
      _unk2              = new String(Enumerable.Repeat('\0', 12).ToArray());

      _scrambledChecksum = 0;
      _fill              = "";
      _solution          = "";
      _clues             = new List<string>();
      _notes             = "";
      _extensions        = new Dictionary<string, string>();
      _puzzleType        = PuzzleType.Normal;
      _solutionState     = SolutionState.Unlocked;
      _helpers           = new List<string>();  // add-ons like Rebus and Markup
    }

    /// <summary>
    /// Read .puz file data and return the Puzzle object.
    /// throws PuzzleFormatError if there's any problem with the file format.
    /// </summary>
    public void Load(byte[] puzData)
    {
      var s = new PuzzleBuffer(puzData);

      // advance to start - files may contain some data before the
      // start of the puzzle use the ACROSS&DOWN magic string as a waypoint
      // save the preamble for round-tripping
      if (!s.SeekTo(Constants.ACROSSDOWN, -2))
        throw new PuzzleFormatError("Data does not appear to represent a puzzle.");


      var puzzle_data = s.Unpack(Constants.HEADER_FORMAT);

      ushort cksum_gbl          = (ushort)       puzzle_data[ 0];
      string acrossDown         = (string)       puzzle_data[ 1];
      ushort cksum_hdr          = (ushort)       puzzle_data[ 2];
      ulong  cksum_magic        = (ulong)        puzzle_data[ 3];
             _fileVersion       = (string)       puzzle_data[ 4];
             _unk1              = (string)       puzzle_data[ 5];
             _scrambledChecksum = (ushort)       puzzle_data[ 6];
             _unk2              = (string)       puzzle_data[ 7];
             _width             = (byte)         puzzle_data[ 8];
             _height            = (byte)         puzzle_data[ 9];
      ushort numclues           = (ushort)       puzzle_data[10];
             _puzzleType        = (PuzzleType)   (ushort)puzzle_data[11];
             _solutionState     = (SolutionState)(ushort)puzzle_data[12];

      _version  = _fileVersion.Substring(_fileVersion.Length - 3);
      _solution = s.Read(_width * _height);
      _fill     = s.Read(_width * _height);

      _title     = s.ReadString();
      _author    = s.ReadString();
      _copyright = s.ReadString();

      for (int clueIndex = 0; clueIndex < numclues; clueIndex++)
        _clues.Add(s.ReadString());

      _notes = s.ReadString();

      var ext_cksum = new Dictionary<string, ushort>();

      while (s.CanUnpack(Constants.EXTENSION_HEADER_FORMAT))
      {
         object[] extHeader = s.Unpack(Constants.EXTENSION_HEADER_FORMAT);

         string code   = (string)extHeader[0];
         ushort length = (ushort)extHeader[1];
         ushort cksum  = (ushort)extHeader[2];

         // extension data is represented as a null-terminated string,
         // but since the data can contain nulls we can't use read_string
         ext_cksum[code]   = cksum;
         _extensions[code] = s.Read(length);

         // extensions have a trailing byte
         s.Read(1);
      }

      // sometimes there's some extra garbage at
      // the end of the file, usually \r\n
      if (s.CanRead())
        _postscript = s.ReadToEnd();

      Debug.WriteLine($"cksum_gbl: {cksum_gbl} calculated {GetGlobalChecksum()}");
      Debug.WriteLine($"cksum_hdr: {cksum_hdr} calculated {GetHeaderChecksum()}");
      Debug.WriteLine($"cksum_magic: {cksum_magic} calculated {GetMagicChecksum()}");

      foreach (var extChecksumPair in ext_cksum)
        Debug.WriteLine($"Ext checksum for {extChecksumPair.Key}: {extChecksumPair.Value} calculated {data_cksum(_encoding.GetBytes(_extensions[extChecksumPair.Key]))}");
    }

    private ushort GetHeaderChecksum()
    {
      return data_cksum(StructConverter.Pack(new object[]
                                             {
                                               _width,
                                               _height,
                                               (ushort)_clues.Count,
                                               (ushort)_puzzleType,
                                               (ushort)_solutionState
                                             }));
    }

    private readonly Encoding _encoding = Encoding.GetEncoding(Constants.DEFAULT_ENCODING);

    private ushort GetTextChecksum(ushort initialChecksum = (ushort)0)
    {
      ushort cksum = initialChecksum;

      // for the checksum to work these fields must be added in order with
      // null termination, followed by all non-empty clues without null
      // termination, followed by notes (but only for version 1.3)
      if (_title != null)
        cksum = data_cksum(_encoding.GetBytes(_title + "\0"), cksum);

      if (_author != null)
          cksum = data_cksum(_encoding.GetBytes(_author + "\0"), cksum);

      if (_copyright != null)
          cksum = data_cksum(_encoding.GetBytes(_copyright + "\0"), cksum);

      foreach (string clue in _clues)
          if (clue != null)
              cksum = data_cksum(_encoding.GetBytes(clue), cksum); // include null terminator?

      // notes included in global cksum only in v1.3 of format
      if (_version == "1.3" && _notes != null && !String.IsNullOrEmpty(_notes))
          cksum = data_cksum(_encoding.GetBytes(_notes), cksum);

       return cksum;
     }

    private ushort GetGlobalChecksum()
    {
      ushort cksum = (ushort)0u;

      cksum = GetHeaderChecksum();
      cksum = data_cksum(_encoding.GetBytes(_solution), cksum);
      cksum = data_cksum(_encoding.GetBytes(_fill), cksum);
      cksum = GetTextChecksum(cksum);

      // extensions do not seem to be included in global cksum
      return cksum;
    }

    private ulong GetMagicChecksum()
    {
      ushort[] cksums =
      {
        GetTextChecksum(),
        data_cksum(_encoding.GetBytes(_fill)),
        data_cksum(_encoding.GetBytes(_solution)),
        GetHeaderChecksum()
      };

      ulong cksum_magic = 0;

      for (int i = 0; i < 4; i++)
      {
        ushort cksum = cksums[i];

        cksum_magic <<= 8;
        cksum_magic |=
        (
          ((ulong)(Constants.MASKSTRING[cksums.Length - i - 1])) ^ (ulong)(cksum & (ushort)0x00ff)
        );

          cksum_magic |=
          (
            (((ulong)Constants.MASKSTRING[cksums.Length - i - 1 + 4]) ^ (ulong)(cksum >> 8)) << 32
          );
      }

      return cksum_magic;
    }

    private ushort data_cksum(byte[] data, ushort initialChecksum = (ushort)0)
    {
      ushort cksum = initialChecksum;

      foreach (byte b in data)
      {
        // right-shift one with wrap-around
        ushort lowbit = (ushort)(cksum & (ushort)1);
        cksum         = (ushort)(cksum >> 1);

        if (lowbit == 1)
          cksum = (ushort)(cksum | 0x8000);

        // then add in the data and clear any carried bit past 16
        cksum = (ushort)((cksum + b) & 0xffff);
      }

      return cksum;
    }
  }

  /// <summary>
  /// Indicates a format error in the .puz file. May be thrown due to
  /// invalid headers, invalid checksum validation, or other format issues.
  /// </summary>
  public class PuzzleFormatError : Exception
  {
    public PuzzleFormatError(string message, Exception inner = null) : base(message, inner) { }
  }

/*
    def save(self, filename):
        with open(filename, 'wb') as f:
            f.write(_tobytes())

    def tobytes(self):
        s = PuzzleBuffer()
        // commit any changes from helpers
        for h in _helpers.values():
            if 'save' in dir(h):
                h.save()

        // include any preamble text we might have found on read
        s.write(_preamble)

        s.pack(HEADER_FORMAT,
               _global_cksum(), ACROSSDOWN.encode(ENCODING),
               _header_cksum(), _magic_cksum(),
               _fileversion, _unk1, _scrambled_cksum,
               _unk2, _width, _height,
               len(_clues), _puzzletype, _solution_state)

        s.write(_solution.encode(ENCODING))
        s.write(_fill.encode(ENCODING))

        s.write_string(_title)
        s.write_string(_author)
        s.write_string(_copyright)

        for clue in _clues:
            s.write_string(clue)

        s.write_string(_notes)

        // do a bit of extra work here to ensure extensions round-trip in the
        // order they were read. this makes verification easier. But allow
        // for the possibility that extensions were added or removed from
        // _extensions
        ext = dict(_extensions)
        for code in __extensions_order:
            data = ext.pop(code, None)
            if data:
                s.pack(EXTENSION_HEADER_FORMAT, code,
                       len(data), data_cksum(data))
                s.write(data + b'\0')

        for code, data in ext.items():
            s.pack(EXTENSION_HEADER_FORMAT, code, len(data), data_cksum(data))
            s.write(data + b'\0')

        s.write(_postscript.encode(ENCODING))

        return s.tobytes()

    def has_rebus(self):
        return _rebus().has_rebus()

    def rebus(self):
        return _helpers.setdefault('rebus', Rebus(self))

    def has_markup(self):
        return _markup().has_markup()

    def markup(self):
        return _helpers.setdefault('markup', Markup(self))

    def clue_numbering(self):
        numbering = DefaultClueNumbering(_fill, _clues,
                                         _width, _height)
        return _helpers.setdefault('clues', numbering)

    def is_solution_locked(self):
        return bool(_solution_state != SolutionState.Unlocked)

    def unlock_solution(self, key):
        if _is_solution_locked():
            unscrambled = unscramble_solution(_solution,
                                              _width, _height, key)
            if not _check_answers(unscrambled):
                return False

            // clear the scrambled bit and cksum
            _solution = unscrambled
            _scrambled_cksum = 0
            _solution_state = SolutionState.Unlocked

        return True

    def lock_solution(self, key):
        if not _is_solution_locked():
            // set the scrambled bit and cksum
            _scrambled_cksum = scrambled_cksum(_solution,
                                                   _width, _height)
            _solution_state = SolutionState.Locked
            scrambled = scramble_solution(_solution,
                                          _width, _height, key)
            _solution = scrambled

    def check_answers(self, fill):
        if _is_solution_locked():
            scrambled = scrambled_cksum(fill, _width, _height)
            return scrambled == _scrambled_cksum
        else:
            return fill == _solution



*/

// clue numbering helper

public enum Direction
{
  Across,
  Down
}

public class ClueNumber
{
  public Direction Direction;
  public int       Number;
  public string    Clue;
  public int       CellIndex;
  public int       Col;
  public int       Row;
  public int       Length;
}

public class DefaultClueNumbering
{
  private string _grid;
  private List<string> _clues;
  private int _width;
  private int _height;

  public List<ClueNumber> Across = new List<ClueNumber>();
  public List<ClueNumber> Down = new List<ClueNumber>();

  public DefaultClueNumbering(string grid, List<string> clues, int width, int height)
  {
      _grid = grid;
      _clues = clues;
      _width = width;
      _height = height;

      // compute across & down
      int clueIndex = 0;
      int clueNumber = 1;

      int lastClueIndex = 0;

      for (int gridIndex = 0; gridIndex < _grid.Length; gridIndex++)
      {
        if (_grid[gridIndex] != '.') // not black square. move to const.
        {
          lastClueIndex = clueIndex;

          bool isAcross = GetCol(gridIndex) == 0 || grid[gridIndex - 1] == '.';

          if (isAcross && GetLengthAcross(gridIndex) > 1)
          {
            Across.Add(new ClueNumber
            {
              Direction = Direction.Across,
              Number = clueNumber,
              Clue = _clues[clueIndex],
              Col = GetCol(gridIndex),
              Row = GetRow(gridIndex),
              CellIndex = gridIndex,
              Length = GetLengthAcross(gridIndex)
            });
            clueIndex++;
          }

          bool isDown = GetRow(gridIndex) == 0 || grid[gridIndex - _width] == '.';

          if (isDown && GetLengthDown(gridIndex) > 1)
          {
            Down.Add(new ClueNumber
            {
              Direction = Direction.Down,
              Number = clueNumber,
              Clue =  _clues[clueIndex],
              Col = GetCol(gridIndex),
              Row = GetRow(gridIndex),
              CellIndex = gridIndex,
              Length = GetLengthDown(gridIndex)
            });

            clueIndex++;
          }

          if (clueIndex > lastClueIndex)
              clueNumber++;
        }
      }
    }

    private int GetCol(int index) => index % _width;
    private int GetRow(int index) => (int)(index / _width);

    private int GetLengthAcross(int index)
    {
      int col = 0;

      //for c in range(0, _width - _col(index)):
      for (col = 0; col < _width - GetCol(index); col++)
      {
        if (_grid[index + col] == '.')
            return col;
      }

      return col + 1;
    }

    private int GetLengthDown(int index)
    {
      int row = 0;

        //for c in range(0, _height - _row(index)):
        for (row = 0; row < _height - GetRow(index); row++)
        {
            if (_grid[index + row * _width] == '.')
                return row;
        }

        return row + 1;
    }
}

/*

class Rebus:
    def __init__(self, puzzle):
        _puzzle = puzzle
        // parse rebus data
        rebus_data = _puzzle.extensions.get(Extensions.Rebus, b'')
        _table = parse_bytes(rebus_data)
        r_sol_data = _puzzle.extensions.get(Extensions.RebusSolutions, b'')
        solutions_str = r_sol_data.decode(ENCODING)
        fill_data = _puzzle.extensions.get(Extensions.RebusFill, b'')
        fill_str = fill_data.decode(ENCODING)
        _solutions = {
            int(item[0]): item[1]
            for item in parse_dict(solutions_str).items()
        }
        _fill = {
            int(item[0]): item[1]
            for item in parse_dict(fill_str).items()
        }

    def has_rebus(self):
        return Extensions.Rebus in _puzzle.extensions

    def is_rebus_square(self, index):
        return bool(_table[index])

    def get_rebus_squares(self):
        return [i for i, b in enumerate(_table) if b]

    def get_rebus_solution(self, index):
        if _is_rebus_square(index):
            return _solutions[_table[index] - 1]
        return None

    def get_rebus_fill(self, index):
        if _is_rebus_square(index):
            return _fill[_table[index] - 1]
        return None

    def set_rebus_fill(self, index, value):
        if _is_rebus_square(index):
            _fill[_table[index] - 1] = value

    def save(self):
        if _has_rebus():
            // commit changes back to puzzle.extensions
            _puzzle.extensions[Extensions.Rebus] = pack_bytes(_table)
            rebus_solutions = dict_to_string(_solutions).encode(ENCODING)
            _puzzle.extensions[Extensions.RebusSolutions] = rebus_solutions
            rebus_fill = dict_to_string(_fill).encode(ENCODING)
            _puzzle.extensions[Extensions.RebusFill] = rebus_fill


class Markup:
    def __init__(self, puzzle):
        _puzzle = puzzle
        // parse markup data
        markup_data = _puzzle.extensions.get(Extensions.Markup, b'')
        _markup = parse_bytes(markup_data)

    def has_markup(self):
        return any(bool(b) for b in _markup)

    def get_markup_squares(self):
        return [i for i, b in enumerate(_markup) if b]

    def is_markup_square(self, index):
        return bool(_table[index])

    def save(self):
        if _has_markup():
            _puzzle.extensions[Extensions.Markup] = pack_bytes(_markup)





def scramble_solution(solution, width, height, key):
    sq = square(solution, width, height)
    data = restore(sq, scramble_string(sq.replace(BLACKSQUARE, ''), key))
    return square(data, height, width)


def scramble_string(s, key):
    """
    s is the puzzle's solution in column-major order, omitting black squares:
    i.e. if the puzzle is:
        C A T
        // // A
        // // R
    solution is CATAR


    Key is a 4-digit number in the range 1000 <= key <= 9999

    """
    key = key_digits(key)
    for k in key:          // foreach digit in the key
        s = shift(s, key)  // for each char by each digit in the key in sequence
        s = s[k:] + s[:k]  // cut the sequence around the key digit
        s = shuffle(s)     // do a 1:1 shuffle of the 'deck'

    return s


def unscramble_solution(scrambled, width, height, key):
    // width and height are reversed here
    sq = square(scrambled, width, height)
    data = restore(sq, unscramble_string(sq.replace(BLACKSQUARE, ''), key))
    return square(data, height, width)


def unscramble_string(s, key):
    key = key_digits(key)
    l = len(s)
    for k in key[::-1]:
        s = unshuffle(s)
        s = s[l-k:] + s[:l-k]
        s = unshift(s, key)

    return s


def scrambled_cksum(scrambled, width, height):
    data = square(scrambled, width, height).replace(BLACKSQUARE, '')
    return data_cksum(data.encode(ENCODING))


def key_digits(key):
    return [int(c) for c in str(key).zfill(4)]


def square(data, w, h):
    aa = [data[i:i+w] for i in range(0, len(data), w)]
    return ''.join(
        [''.join([aa[r][c] for r in range(0, h)]) for c in range(0, w)]
    )


def shift(s, key):
    atoz = string.ascii_uppercase
    return ''.join(
        atoz[(atoz.index(c) + key[i % len(key)]) % len(atoz)]
        for i, c in enumerate(s)
    )


def unshift(s, key):
    return shift(s, [-k for k in key])


def shuffle(s):
    mid = int(math.floor(len(s) / 2))
    items = functools.reduce(operator.add, zip(s[mid:], s[:mid]))
    return ''.join(items) + (s[-1] if len(s) % 2 else '')


def unshuffle(s):
    return s[1::2] + s[::2]


def restore(s, t):
    """
    s is the source string, it can contain '.'
    t is the target, it's smaller than s by the number of '.'s in s

    Each char in s is replaced by the corresponding
    char in t, jumping over '.'s in s.

    >>> restore('ABC.DEF', 'XYZABC')
    'XYZ.ABC'
    """
    t = (c for c in t)
    return ''.join(next(t) if not is_blacksquare(c) else c for c in s)


def is_blacksquare(c):
    if isinstance(c, int):
        c = chr(c)
    return c == BLACKSQUARE
*/


/*
#
// functions for parsing / serializing primitives
#


def parse_bytes(s):
    return list(struct.unpack('B' * len(s), s))


def pack_bytes(a):
    return struct.pack('B' * len(a), *a)


// dict string format is k1:v1;k2:v2;...;kn:vn;
// (for whatever reason there's a trailing ';')
def parse_dict(s):
    return dict(p.split(':') for p in s.split(';') if ':' in p)


def dict_to_string(d):
    return ';'.join(':'.join(map(str, [k, v])) for k, v in d.items()) + ';'
*/

}
