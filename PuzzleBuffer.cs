using System;
using System.Text;

namespace PuzXls
{
  /// <summary>
  /// wraps a data buffer and provides .puz-specific methods for reading and writing data
  /// </summary>
  class PuzzleBuffer
  {
    private readonly Encoding _encoding;

    private readonly byte[] _data;
    private          int    _pos;

    public PuzzleBuffer(byte[] puzData = null, string encoding = Constants.DEFAULT_ENCODING)
    {
      _data     = puzData ?? new byte[] { };
      _encoding = Encoding.GetEncoding(encoding);
      _pos      = 0;
    }

    public bool CanRead(int numBytes = 1) => _pos + numBytes <= _data.Length;

    public int Length => _data.Length;

    public string Read(int bytes)
    {
      int start = _pos;
      _pos += bytes;

      System.Diagnostics.Debug.Write($"Read {bytes} bytes: ");
      byte[] subArray = new byte[bytes];
      Array.Copy(_data, start, subArray, 0, bytes);

      string content = _encoding.GetString(subArray);
      System.Diagnostics.Debug.WriteLine($" '{content}'");
      return content;
    }

    public string ReadToEnd()
    {
      return Read(_data.Length - _pos);
    }

    public string ReadString()
    {
      return ReadUntil(0);
    }

    public string ReadUntil(byte terminator)
    {
      int end   = Array.IndexOf(_data, terminator, _pos);
      int bytes = (end - _pos);

      var stringValue = bytes == 0 ? String.Empty : Read(bytes);

      _pos++; // skip the null-term character but don't include it

      return stringValue;
    }

    public bool SeekTo(string content, int offset = 0)
    {
      try
      {
        _pos = _data.Locate(_encoding.GetBytes(content)) + offset;
        return true;
      }
      catch (ValueErrorException)
      {
        // Not found, advance to end.
        _pos = Length;
        return false;
      }
    }

    public bool CanUnpack(string structFormat)
    {
      return CanRead(StructConverter.CalcSize(structFormat));
    }

    public object[] Unpack(string structFormat)
    {
      int start = _pos;

      try
      {
        var res = StructConverter.Unpack(structFormat, _data, _pos);
        _pos += StructConverter.CalcSize(structFormat);

        return res;
      }
      catch (Exception exception)
      {
        throw new PuzzleFormatError($"could not unpack values at {start} for format {structFormat}", exception);
      }
    }

        //def write(self, s):
    //    _data.append(s)

    //def write_string(self, s):
    //    s = s or ''
    //    _data.append(s.encode(ENCODING) + b'\0')

    //def pack(self, struct_format, *values):
    //    _data.append(struct.pack(struct_format, *values))
  }
}
