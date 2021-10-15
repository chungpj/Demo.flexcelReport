using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Report.Common
{
    public interface IChecksum
    {
        UInt64 Checksum { get; }

        UInt64 Update(byte[] buffer);
        UInt64 Update(byte[] buffer, int offset, int count);
        void Reset();
    };

    public static class StreamHelper
    {
        private const int STREAM_BUFFER_SIZE = 32768;

        public static System.Drawing.Image LoadImage(Stream stream)
        {
            var image = System.Drawing.Image.FromStream(stream);
            stream.Seek(0, SeekOrigin.Begin);
            image.Tag = StreamHelper.Read(stream);
            return image;
        }

        public static System.Drawing.Image LoadImage(byte[] data)
        {
            using (var stream = new MemoryStream(data))
            {
                var image = System.Drawing.Image.FromStream(stream);
                stream.Seek(0, SeekOrigin.Begin);
                image.Tag = data;
                return image;
            }
        }

        public static System.Drawing.Image InitStream(System.Drawing.Image image)
        {
            image.Tag = StreamHelper.Read(image);
            return image;
        }

        public static byte[] Read(System.Drawing.Image image)
        {
            using (var stream = new MemoryStream())
            {
                image.Save(stream, image.RawFormat);
                stream.Seek(0, SeekOrigin.Begin);
                return StreamHelper.Read(stream);
            }
        }

        public static byte[] Read(System.Drawing.Icon icon)
        {
            using (var stream = new MemoryStream())
            {
                icon.Save(stream);
                stream.Seek(0, SeekOrigin.Begin);
                return StreamHelper.Read(stream);
            }
        }

        public static byte[] Read(Stream stream)
        {
            using (stream)
            using (BinaryReader reader = new BinaryReader(stream))
            {
                byte[] data = new byte[stream.Length];
                if (reader.Read(data, 0, data.Length) != data.Length)
                    throw new Exception("InternalBufferOverflowException");
                return data;
            }
        }

        public static void WriteUtf8(this Stream stream, string utf8String)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(utf8String);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static void WriteUtf8Format(this Stream stream, string utf8Format, params object[] args)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(String.Format(utf8Format, args));
            stream.Write(bytes, 0, bytes.Length);
        }

        public static void WriteUtf8LineFormat(this Stream stream, string utf8Format, params object[] args)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(String.Format(utf8Format, args));
            stream.Write(bytes, 0, bytes.Length);
            bytes = Encoding.UTF8.GetBytes(Environment.NewLine);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static void WriteUtf8Line(this Stream stream)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(Environment.NewLine);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static void WriteUtf8Line(this Stream stream, string utf8String)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(utf8String);
            stream.Write(bytes, 0, bytes.Length);
            bytes = Encoding.UTF8.GetBytes(Environment.NewLine);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static int Streaming(Stream source, Stream destination)
        {
            byte[] buff = new byte[STREAM_BUFFER_SIZE];
            int len, total = 0;
            while ((len = source.Read(buff, 0, buff.Length)) > 0)
            {
                destination.Write(buff, 0, len);
                total += len;
            }
            return total;
        }

        public static long StreamCopy(Stream source, Stream destination)
        {
            return StreamCopy(source, destination, source.Length - source.Position, null);
        }

        public static long StreamCopy(Stream source, Stream destination, IChecksum checksum)
        {
            return StreamCopy(source, destination, source.Length - source.Position, checksum);
        }

        public static long StreamCopy(Stream source, Stream destination, long length)
        {
            return StreamCopy(source, destination, length, null);
        }

        public static long StreamCopy(Stream source, Stream destination, long length, IChecksum checksum)
        {
            if (null == source)
            {
                throw new ArgumentNullException("source");
            }

            if (null == destination)
            {
                throw new ArgumentNullException("destination");
            }

            if (length < 0)
            {
                throw new ArgumentOutOfRangeException("length");
            }

            byte[] buff = new byte[STREAM_BUFFER_SIZE];

            int bytesCopied = 0;
            while (length > 0)
            {
                int toRead = Math.Min(STREAM_BUFFER_SIZE, (int)length);
                int bytesRead = source.Read(buff, 0, toRead);

                if (bytesRead == 0)
                    break;

                if (null != checksum)
                    checksum.Update(buff, 0, bytesRead);

                destination.Write(buff, 0, bytesRead);

                length -= bytesRead;
                bytesCopied += bytesRead;
            }

            return bytesCopied;
        }

        public static long StreamCopy(Stream source, long first, Stream destination)
        {
            return StreamCopy(source, first, destination, destination.Position, source.Length - first, null);
        }

        public static long StreamCopy(Stream source, long first, Stream destination, IChecksum checksum)
        {
            return StreamCopy(source, first, destination, destination.Position, source.Length - first, checksum);
        }

        public static long StreamCopy(Stream source, long first, Stream destination, long result)
        {
            return StreamCopy(source, first, destination, result, source.Length - first, null);
        }

        public static long StreamCopy(Stream source, long first, Stream destination, long result, long length)
        {
            return StreamCopy(source, first, destination, result, length, null);
        }

        public static long StreamCopy(Stream source, long first, Stream destination, long result, long length, IChecksum checksum)
        {
            if (null == source)
            {
                throw new ArgumentNullException("source");
            }

            if (null == destination)
            {
                throw new ArgumentNullException("destination");
            }

            if (first < 0)
            {
                throw new ArgumentOutOfRangeException("first");
            }

            if (result < 0)
            {
                throw new ArgumentOutOfRangeException("result");
            }

            if (length < 0)
            {
                throw new ArgumentOutOfRangeException("length");
            }

            byte[] buff = new byte[STREAM_BUFFER_SIZE];

            int bytesCopied = 0;
            while (length > 0)
            {
                int toRead = Math.Min(STREAM_BUFFER_SIZE, (int)length);
                source.Seek(first + bytesCopied, SeekOrigin.Begin);
                int bytesRead = source.Read(buff, 0, toRead);

                if (bytesRead == 0)
                    break;

                if (checksum != null)
                    checksum.Update(buff, 0, bytesRead);

                destination.Seek(result + bytesCopied, SeekOrigin.Begin);
                destination.Write(buff, 0, bytesRead);

                length -= bytesRead;
                bytesCopied += bytesRead;
            }

            return bytesCopied;
        }

        public static long StreamChecksum(Stream source, IChecksum checksum)
        {
            return StreamChecksum(source, checksum, source.Length - source.Position);
        }

        public static long StreamChecksum(Stream source, IChecksum checksum, long length)
        {
            if (null == source)
            {
                throw new ArgumentNullException("source");
            }

            if (null == checksum)
            {
                throw new ArgumentNullException("checksum");
            }

            if (length < 0)
            {
                throw new ArgumentOutOfRangeException("length");
            }

            byte[] buff = new byte[STREAM_BUFFER_SIZE];

            int bytesCopied = 0;
            while (length > 0)
            {
                int toRead = Math.Min(STREAM_BUFFER_SIZE, (int)length);
                int bytesRead = source.Read(buff, 0, toRead);

                if (bytesRead == 0)
                    break;

                checksum.Update(buff, 0, bytesRead);

                length -= bytesRead;
                bytesCopied += bytesRead;
            }

            return bytesCopied;
        }

        public static Stream Zip(string path)
        {
            return StreamHelper.Zip(path, null);
        }

        public static Stream Zip(string path, Stream output)
        {
            using (var input = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                //Prepare for compress
                bool autoOutput = (output == null);
                if (autoOutput)
                    output = new MemoryStream();

                using (var gzip = new GZipStream(output, CompressionMode.Compress, true))
                {
                    //Compress
                    input.CopyTo(gzip);

                    //Close, DO NOT FLUSH cause bytes will go missing...
                    gzip.Close();

                    if (autoOutput)
                        output.Seek(0, SeekOrigin.Begin);

                    return output;
                }
            }
        }

        public static void UnZip(Stream input, string path)
        {
            using (input)
            {
                //Prepare for decompress
                using (var output = File.Open(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
                using (var gzip = new GZipStream(input, CompressionMode.Decompress))
                {
                    //Compress
                    gzip.CopyTo(output);

                    //Close, DO NOT FLUSH cause bytes will go missing...
                    gzip.Close();
                }
            }
        }

        public static Stream Deflate(string path)
        {
            return StreamHelper.Deflate(path, null);
        }

        public static Stream Deflate(string path, Stream output)
        {
            using (var input = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                //Prepare for compress
                bool autoOutput = (output == null);
                if (autoOutput)
                    output = new MemoryStream();

                using (var gzip = new DeflateStream(output, CompressionMode.Compress, true))
                {
                    //Compress
                    input.CopyTo(gzip);

                    //Close, DO NOT FLUSH cause bytes will go missing...
                    gzip.Close();

                    if (autoOutput)
                        output.Seek(0, SeekOrigin.Begin);

                    return output;
                }
            }
        }

        public static void UnDeflate(Stream input, string path)
        {
            using (input)
            {
                //Prepare for decompress
                using (var output = File.Open(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
                using (var gzip = new DeflateStream(input, CompressionMode.Decompress))
                {
                    //Compress
                    gzip.CopyTo(output);

                    //Close, DO NOT FLUSH cause bytes will go missing...
                    gzip.Close();
                }
            }
        }

    }
}
