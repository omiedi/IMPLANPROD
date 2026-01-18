using System;
using System.Text;

public class MbfAnalyzer
{
    public static void Main()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        float[] values = { 15f, 60f, 6f, 12f, 10f, 20f };
        
        Console.WriteLine("Value | MBF Bytes (Hex) | MBF String (Latin1)");
        Console.WriteLine("---|---|---");
        
        foreach (var val in values)
        {
            byte[] mbf = FloatToMbf(val);
            string hex = BitConverter.ToString(mbf);
            string str = Encoding.GetEncoding(1252).GetString(mbf);
            // Replace non-printable chars for display
            string safeStr = "";
            foreach(char c in str)
            {
               if (char.IsControl(c)) safeStr += ".";
               else safeStr += c;
            }
            
            Console.WriteLine($"{val} | {hex} | {safeStr}");
        }
    }

    private static byte[] FloatToMbf(float val)
    {
        if (val == 0.0f) return new byte[] { 0, 0, 0, 0 };

        // IEEE to MBF conversion (Reverse of the Read logic)
        // This is an approximation since we usually only read.
        // But for these small integers, we can likely deduce the MBF pattern manually or inverted.
        
        // Actually, easier to construct MBF manually for simple integers if we know the format.
        // MBF Format (4 bytes):
        // Byte 0: Mantissa (LSB)
        // Byte 1: Mantissa
        // Byte 2: Sign(1) + Mantissa (7)
        // Byte 3: Exponent (Excess-128? Code says aa4 - 129, so Exponent + 129?)
        
        // Let's rely on the fact that existing code *reads* it. 
        // We can brute force "Find float X that maps to MBF bytes" isn't easy without the Write function.
        // But we can assume standard MBF. 
        // Let's try to just use the Microsoft.Basic.VisualBasic or similar? No, standard .NET.
        
        // Let's implement a quick "Double to MBF" converter.
        // https://en.wikipedia.org/wiki/Microsoft_Binary_Format
        // 4-byte MBF:
        // Exp: 8 bits (offset 3)
        // Sign: 1 bit (in offset 2)
        // Mantissa: 23 bits (rest of 2, 1, 0)
        
        // Value = (1 + Mantissa) * 2^(Exp-129)
        
        double d = (double)val;
        if (d == 0) return new byte[4];
        
        int sign = d < 0 ? -1 : 1;
        d = Math.Abs(d);
        
        int exponent = (int)Math.Floor(Math.Log(d) / Math.Log(2));
        double mantissa = d / Math.Pow(2, exponent) - 1.0;
        
        int mbfExponent = exponent + 129;
        
        // Mantissa is 23 bits.
        int mantissaInt = (int)(mantissa * 8388608.0 + 0.5); // 2^23
        
        byte[] bytes = new byte[4];
        bytes[3] = (byte)mbfExponent;
        bytes[2] = (byte)((mantissaInt >> 16) & 0x7F);
        if (sign < 0) bytes[2] |= 0x80;
        
        bytes[1] = (byte)((mantissaInt >> 8) & 0xFF);
        bytes[0] = (byte)(mantissaInt & 0xFF);
        
        return bytes;
    }
}
