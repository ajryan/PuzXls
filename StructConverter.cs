using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;

namespace PuzXls
{
  public class StructConverter
  {
    private static readonly Encoding _encoding = Encoding.GetEncoding(Constants.DEFAULT_ENCODING);

    // We use this function to provide an easier way to type-agnostically call the GetBytes method of the BitConverter class.
    // This means we can have much cleaner code below.
    private static byte[] TypeAgnosticGetBytes(object o)
    {
      if (o is int)                return BitConverter.GetBytes((int)o);
      if (o is uint)               return BitConverter.GetBytes((uint)o);
      if (o is long)               return BitConverter.GetBytes((long)o);
      if (o is ulong)              return BitConverter.GetBytes((ulong)o);
      if (o is short)              return BitConverter.GetBytes((short)o);
      if (o is ushort)             return BitConverter.GetBytes((ushort)o);
      if (o is string)             return _encoding.GetBytes((string)o); 
      if (o is byte || o is sbyte) return new byte[] { (byte)o };

      throw new ArgumentException("Unsupported object type found");
    }

    private static string GetFormatSpecifierFor(object o)
    {
      if (o is int)    return "l";
      if (o is uint)   return "L";
      if (o is long)   return "q";
      if (o is ulong)  return "Q";
      if (o is short)  return "h";
      if (o is ushort) return "H";
      if (o is byte)   return "B";
      if (o is sbyte)  return "b";
      if (o is string) return $"{((string)o).Length}s";

      throw new ArgumentException("Unsupported object type found");
    }

    /// <summary>
    /// Convert a byte array into an array of objects based on Python's "struct.unpack" protocol.
    /// </summary>
    /// <param name="fmt">A "struct.pack"-compatible format string</param>
    /// <param name="bytes">An array of bytes to convert to objects</param>
    /// <returns>Array of objects.</returns>
    /// <remarks>You are responsible for casting the objects in the array back to their proper types.</remarks>
    public static object[] Unpack(string fmt, byte[] bytes, int startPos = 0)
    {
      Debug.WriteLine($"Format string '{fmt}' is length {fmt.Length}, {bytes.Length} bytes provided.");

      // First we parse the format string to make sure it's proper.
      if (fmt.Length < 1) throw new ArgumentException("Format string cannot be empty.");

      bool endianFlip = false;

      if (fmt.Substring(0, 1) == "<")
      {
        Debug.WriteLine("  Endian marker found: little endian");
        // Little endian.
        // Do we need to flip endianness?
        if (BitConverter.IsLittleEndian == false) endianFlip = true;
        fmt = fmt.Substring(1);
      }
      else if (fmt.Substring(0, 1) == ">")
      {
        Debug.WriteLine("  Endian marker found: big endian");
        // Big endian.
        // Do we need to flip endianness?
        if (BitConverter.IsLittleEndian == true) endianFlip = true;
        fmt = fmt.Substring(1);
      }

      // Now, we find out how long the byte array needs to be
      int totalByteLength = CalcSize(fmt);

      Debug.WriteLine($"Endianness will {(object)(endianFlip ? "" : "NOT ")}be flipped.");
      Debug.WriteLine($"The byte array is expected to be {totalByteLength} bytes long.");

      // Test the byte array length to see if it contains as many bytes as is needed for the string.
      if ((bytes.Length - startPos) < totalByteLength)
        throw new ArgumentException("The number of bytes provided does not match the total length of the format string.");

      // Ok, we can go ahead and start parsing bytes!
      int byteArrayPosition = startPos;
      List<object> outputList = new List<object>();

      Debug.WriteLine("Processing byte array...");
      string currentStringSize = null;
      foreach (char c in fmt)
      {
        byte[] buf;

        switch (c)
        {
          case 'q':
            outputList.Add(BitConverter.ToInt64(bytes, byteArrayPosition));
            byteArrayPosition+=8;
            Debug.Write("  q: Added signed 64-bit integer");
            break;
          case 'Q':
            outputList.Add(BitConverter.ToUInt64(bytes, byteArrayPosition));
            byteArrayPosition+=8;
            Debug.Write("  Q: Added unsigned 64-bit integer");
            break;
          case 'l':
            outputList.Add(BitConverter.ToInt32(bytes, byteArrayPosition));
            byteArrayPosition+=4;
            Debug.Write("  l: Added signed 32-bit integer");
            break;
          case 'L':
            outputList.Add(BitConverter.ToUInt32(bytes, byteArrayPosition));
            byteArrayPosition+=4;
            Debug.Write("  L: Added unsignedsigned 32-bit integer");
            break;
          case 'h':
            outputList.Add(BitConverter.ToInt16(bytes, byteArrayPosition));
            byteArrayPosition += 2;
            Debug.Write("  h: Added signed 16-bit integer");
            break;
          case 'H':
            outputList.Add(BitConverter.ToUInt16(bytes, byteArrayPosition));
            byteArrayPosition += 2;
            Debug.Write("  H: Added unsigned 16-bit integer");
            break;
          case 'b':
            buf = new byte[1];
            Array.Copy(bytes, byteArrayPosition, buf, 0, 1);
            outputList.Add((sbyte)buf[0]);
            byteArrayPosition++;
            Debug.Write("  b: Added signed byte");
            break;
          case 'B':
            buf = new byte[1];
            Array.Copy(bytes, byteArrayPosition, buf, 0, 1);
            outputList.Add(buf[0]);
            byteArrayPosition++;
            Debug.Write("  B: Added unsigned byte");
            break;
          case 'x':
            byteArrayPosition++;
            Debug.WriteLine("  x: Ignoring a byte");
            break;
          case 's':
            int stringSize = Int32.Parse(currentStringSize);
            buf = new byte[stringSize];
            Array.Copy(bytes, byteArrayPosition, buf, 0, stringSize);
            outputList.Add(_encoding.GetString(buf));
            byteArrayPosition += stringSize;
            Debug.Write($"  {stringSize}s: Added string");
            break;
          default:
            if (c < '0' && c > '9')
              throw new ArgumentException("You should not be here.");
            break;
        }

        if (c >= '0' && c <= '9')
          currentStringSize += c;
        else
        {
          if (outputList.Count > 0 && c != ' ' && c != '\r' && c != '\n' && c != 'x')
            Debug.WriteLine($" '{outputList.Last()}'");

          currentStringSize = null;
        }
      }

      Debug.WriteLine("... done unpack");
      return outputList.ToArray();
    }

    /// <summary>
    /// Convert an array of objects to a byte array, along with a string that can be used with Unpack.
    /// </summary>
    /// <param name="items">An object array of items to convert</param>
    /// <param name="littleEndian">Set to False if you want to use big endian output.</param>
    /// <param name="neededFormatStringToRecover">Variable to place an 'Unpack'-compatible format string into.</param>
    /// <returns>A Byte array containing the objects provided in binary format.</returns>
    public static byte[] Pack(object[] items, bool littleEndian, out string neededFormatStringToRecover)
    {
      // make a byte list to hold the bytes of output
      List<byte> outputBytes = new List<byte>();

      // should we be flipping bits for proper endinanness?
      bool endianFlip = (littleEndian != BitConverter.IsLittleEndian);

      // start working on the output string
      string outString = (littleEndian == false ? ">" : "<");

      // convert each item in the objects to the representative bytes
      foreach (object o in items)
      {
        byte[] theseBytes = TypeAgnosticGetBytes(o);

        if (endianFlip == true)
          theseBytes = (byte[])theseBytes.Reverse().ToArray();

        outString += GetFormatSpecifierFor(o);
        outputBytes.AddRange(theseBytes);
      }

      neededFormatStringToRecover = outString;

      return outputBytes.ToArray();
    }

    public static byte[] Pack(object[] items)
    {
      string dummy = "";
      return Pack(items, true, out dummy);
    }

    public static int CalcSize(string fmt)
    {
      int totalByteLength = 0;

      string currentStringSize = null;

      foreach (char c in fmt)
      {
        if (c != ' ' && c != '\r' && c != '\n')
          Debug.WriteLine($"  Format character found: {c}");

        switch (c)
        {
          case 'q':
          case 'Q':
            totalByteLength += sizeof(ulong);
            break;
          case 'l':
          case 'L':
            totalByteLength += sizeof(uint);
            break;
          case 'h':
          case 'H':
            totalByteLength += sizeof(ushort);
            break;
          case 'b':
          case 'B':
          case 'x':
            totalByteLength += sizeof(byte);
            break;
        }

        if (c >= '0' && c <= '9')
          currentStringSize += c;
        else if (currentStringSize != null)
        {
          int stringSize = Int32.Parse(currentStringSize);
          totalByteLength += stringSize;
          currentStringSize = null;
        }
      }

      return totalByteLength;
    }
  }
}
