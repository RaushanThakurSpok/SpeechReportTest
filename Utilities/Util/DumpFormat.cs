using System;
using System.Collections.Generic;
using System.Text;

namespace Amcom.SDC.CodeUtilities.Util
{
    public static class DumpFormat
    {

        public static string HexDump(byte[] bytes)
        {
            return HexDump(bytes, 16);
        }


        public static string HexDump(byte[] bytes,int outputWidth)
        {
            int width = outputWidth;
            bool done = false;
            int maxbytes = bytes.Length;
            int bytesread = 0;
            int start = 0;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            while (!done)
            {
                int bytesToRead = (maxbytes - bytesread < width ? maxbytes - bytesread : width);
                bytesread += bytesToRead;
                sb.Append(bytesread.ToString("D8") + ": ");
                for (int index = start; index <= (start + bytesToRead) - 1; index++)
                {
                    sb.Append(string.Format("{0:X2} ", bytes[index]));
                }
                if (bytesToRead < width)
                {
                    char c =' ';
                    sb.Append(new string(c, (width - bytesToRead) * 3));
                }

                sb.Append(": ");
                for (int index = start; index <= (start + bytesToRead) - 1; index++)
                {
                    byte b = bytes[index];
                    sb.Append((b > 32 & b < 126 ? char.ConvertFromUtf32(b) : "."));
                }
                sb.Append("\r\n");
                done = bytesread == maxbytes;
                if (done)
                    break; // TODO: might not be correct. Was : Exit While
                start = (bytesread);
            }
            return sb.ToString();
        }
        //public static string HexDump(byte[] bytes)
        //{
        //    if (bytes == null) return "<null>";
        //    int len = bytes.Length;
        //    StringBuilder result = new StringBuilder(((len + 15) / 16) * 78);
        //    char[] chars = new char[78];
        //    // fill all with blanks
        //    for (int i = 0; i < 75; i++) chars[i] = ' ';
        //    chars[76] = '\r';
        //    chars[77] = '\n';

        //    for (int i1 = 0; i1 < len; i1 += 16)
        //    {
        //        chars[0] = HexChar(i1 >> 28);
        //        chars[1] = HexChar(i1 >> 24);
        //        chars[2] = HexChar(i1 >> 20);
        //        chars[3] = HexChar(i1 >> 16);
        //        chars[4] = HexChar(i1 >> 12);
        //        chars[5] = HexChar(i1 >> 8);
        //        chars[6] = HexChar(i1 >> 4);
        //        chars[7] = HexChar(i1 >> 0);

        //        int offset1 = 11;
        //        int offset2 = 60;

        //        for (int i2 = 0; i2 < 16; i2++)
        //        {
        //            if (i1 + i2 >= len)
        //            {
        //                chars[offset1] = ' ';
        //                chars[offset1 + 1] = ' ';
        //                chars[offset2] = ' ';
        //            }
        //            else
        //            {
        //                byte b = bytes[i1 + i2];
        //                chars[offset1] = HexChar(b >> 8);
        //                chars[offset1 + 1] = HexChar(b);
        //                chars[offset2] = (b < 32 ? '·' : (char)b);
        //            }
        //            offset1 += (i2 == 8 ? 4 : 3);
        //            offset2++;
        //        }
        //        result.Append(chars);
        //    }
        //    return result.ToString();
        //}

        private static char HexChar(int value)
        {
            value &= 0xF;
            if (value >= 0 && value <= 9)
                return (char)('0' + value);
            return (char)('A' + (value - 10));
        }
    }
}


