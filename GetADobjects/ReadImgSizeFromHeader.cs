using System;
using System.Collections.Generic;
using System.IO;


public struct ImgSize
{
    public Int32 Width;
    public Int32 Height;
    public ImgSize(Int32 width, Int32 height)
    {
        this.Width = width;
        this.Height = height;
    }
    public bool IsEmpty()
    {
        if (Width == 0 && Height == 0)
            return true;
        return false;
    }
};

// Source: http://www.codeproject.com/Articles/35978/Reading-Image-Headers-to-Get-Width-and-Height
/// <summary>
/// Taken from http://stackoverflow.com/questions/111345/getting-image-dimensions-without-reading-the-entire-file/111349
/// Minor improvements including supporting unsigned 16-bit integers when decoding Jfif and added logic
/// to load the image using new Bitmap if reading the headers fails
/// </summary>
public static class ImageHeader
{
    private static Dictionary<byte[], Func<BinaryReader, ImgSize>> imageFormatDecoders = new Dictionary<byte[], Func<BinaryReader, ImgSize>>()
    { 
        { new byte[] { 0xff, 0xd8 }, DecodeJfif }, 
        { new byte[] { 0x42, 0x4D }, DecodeBitmap }, 
        { new byte[] { 0x47, 0x49, 0x46, 0x38, 0x37, 0x61 }, DecodeGif }, 
        { new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, DecodeGif }, 
        { new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, DecodePng },
    };

    public static ImgSize GetDimensions(byte[] imgdata)
    {
        MemoryStream memstream = null;
        ImgSize imgsize = new ImgSize(0, 0);
        try
        {
            memstream = new MemoryStream(imgdata);

            using (BinaryReader binaryReader = new BinaryReader(memstream))
            {
                imgsize = GetDimensions(binaryReader);
            }
        }
        finally
        {
            if (memstream != null)
                memstream.Close();
        }
        return imgsize;
    }

    /// <summary>        
    /// Gets the dimensions of an image.        
    /// </summary>        
    /// <param name="binaryReader">Binary data of the image to get the dimensions of.</param>        
    /// <returns>The dimensions of the specified image.</returns>        
    public static ImgSize GetDimensions(BinaryReader binaryReader)
    {
        int maxMagicBytesLength = 8;
        byte[] magicBytes = new byte[maxMagicBytesLength];
        for (int i = 0; i < maxMagicBytesLength; i += 1)
        {
            magicBytes[i] = binaryReader.ReadByte();
            foreach (var kvPair in imageFormatDecoders)
            {
                if (StartsWith(magicBytes, kvPair.Key))
                {
                    return kvPair.Value(binaryReader);
                }
            }
        }
        return new ImgSize(0, 0);
    }

    private static bool StartsWith(byte[] thisBytes, byte[] thatBytes)
    {
        for (int i = 0; i < thatBytes.Length; i += 1)
        {
            if (thisBytes[i] != thatBytes[i])
            {
                return false;
            }
        }
        return true;
    }

    private static short ReadLittleEndianInt16(BinaryReader binaryReader)
    {
        byte[] bytes = new byte[sizeof(short)];

        for (int i = 0; i < sizeof(short); i += 1)
        {
            bytes[sizeof(short) - 1 - i] = binaryReader.ReadByte();
        }
        return BitConverter.ToInt16(bytes, 0);
    }

    private static ushort ReadLittleEndianUInt16(BinaryReader binaryReader)
    {
        byte[] bytes = new byte[sizeof(ushort)];

        for (int i = 0; i < sizeof(ushort); i += 1)
        {
            bytes[sizeof(ushort) - 1 - i] = binaryReader.ReadByte();
        }
        return BitConverter.ToUInt16(bytes, 0);
    }

    private static int ReadLittleEndianInt32(BinaryReader binaryReader)
    {
        byte[] bytes = new byte[sizeof(int)];
        for (int i = 0; i < sizeof(int); i += 1)
        {
            bytes[sizeof(int) - 1 - i] = binaryReader.ReadByte();
        }
        return BitConverter.ToInt32(bytes, 0);
    }

    private static ImgSize DecodeBitmap(BinaryReader binaryReader)
    {
        binaryReader.ReadBytes(16);
        int width = binaryReader.ReadInt32();
        int height = binaryReader.ReadInt32();
        return new ImgSize(width, height);
    }

    private static ImgSize DecodeGif(BinaryReader binaryReader)
    {
        int width = binaryReader.ReadInt16();
        int height = binaryReader.ReadInt16();
        return new ImgSize(width, height);
    }

    private static ImgSize DecodePng(BinaryReader binaryReader)
    {
        binaryReader.ReadBytes(8);
        int width = ReadLittleEndianInt32(binaryReader);
        int height = ReadLittleEndianInt32(binaryReader);
        return new ImgSize(width, height);
    }

    private static ImgSize DecodeJfif(BinaryReader binaryReader)
    {
        while (binaryReader.ReadByte() == 0xff)
        {
            byte marker = binaryReader.ReadByte();
            short chunkLength = ReadLittleEndianInt16(binaryReader);
            if (marker == 0xc0)
            {
                binaryReader.ReadByte();
                int height = ReadLittleEndianInt16(binaryReader);
                int width = ReadLittleEndianInt16(binaryReader);
                return new ImgSize(width, height);
            }

            if (chunkLength < 0)
            {
                ushort uchunkLength = (ushort)chunkLength;
                binaryReader.ReadBytes(uchunkLength - 2);
            }
            else
            {
                binaryReader.ReadBytes(chunkLength - 2);
            }
        }
        return new ImgSize(0, 0);
    }
}

