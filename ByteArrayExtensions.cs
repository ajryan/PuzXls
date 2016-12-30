using System;

namespace PuzXls
{
  // TOOD: better name
  public class ValueErrorException : Exception { }

  public static class ByteArrayExtensions
  {
    public static int Locate(this byte[] self, byte[] candidate, int startPos = 0)
    {
      if (IsEmptyLocate(self, candidate))
        throw new ValueErrorException();

      for (int pos = startPos; pos < self.Length; pos++)
      {
        if (!IsMatch(self, candidate, pos))
          continue;

        return pos;
      }

      throw new ValueErrorException();
    }

    private static bool IsMatch(byte[] array, byte[] candidate, int arrayPos)
    {
      if (candidate.Length > (array.Length - arrayPos))
        return false;

      for (int candidatePos = 0; candidatePos < candidate.Length; candidatePos++)
        if (array[arrayPos + candidatePos] != candidate[candidatePos])
          return false;

      return true;
    }

    private static bool IsEmptyLocate(byte[] array, byte[] candidate)
    {
      return array            == null ||
             candidate        == null ||
             array.Length     == 0 ||
             candidate.Length == 0 ||
             candidate.Length > array.Length;
    }
  }
}
